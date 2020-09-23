VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coolest Fader 9.3"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "Frames.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Fade!"
      Height          =   2775
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      Begin VB.PictureBox Picture1a 
         Height          =   135
         Left            =   300
         ScaleHeight     =   75
         ScaleWidth      =   915
         TabIndex        =   59
         Top             =   430
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   3960
         TabIndex        =   7
         Tag             =   "0"
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ListBox List2 
         Height          =   450
         Left            =   4200
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   2160
         Top             =   1080
      End
      Begin VB.PictureBox Picture2 
         Height          =   135
         Left            =   300
         ScaleHeight     =   75
         ScaleWidth      =   915
         TabIndex        =   44
         ToolTipText     =   "Green"
         Top             =   1030
         Width           =   975
      End
      Begin VB.PictureBox Picture3 
         Height          =   135
         Left            =   300
         ScaleHeight     =   75
         ScaleWidth      =   915
         TabIndex        =   43
         ToolTipText     =   "Blue"
         Top             =   1630
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         MaxLength       =   45
         TabIndex        =   18
         Tag             =   "This is a preview of the colors in your selected fade."
         ToolTipText     =   "Message"
         Top             =   2400
         Width           =   3975
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "Frames.frx":0442
         Left            =   1440
         List            =   "Frames.frx":0444
         TabIndex        =   3
         ToolTipText     =   "List of Colors"
         Top             =   720
         Width           =   1215
      End
      Begin MSComctlLib.Slider HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   45
         ToolTipText     =   "Red"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   30
         Max             =   255
         TickFrequency   =   30
      End
      Begin MSComctlLib.Slider HScroll2 
         Height          =   255
         Left            =   240
         TabIndex        =   46
         ToolTipText     =   "Green"
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   30
         Max             =   255
         TickFrequency   =   30
      End
      Begin MSComctlLib.Slider HScroll3 
         Height          =   255
         Left            =   240
         TabIndex        =   47
         ToolTipText     =   "Blue"
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         Max             =   255
         TickFrequency   =   30
      End
      Begin VB.Label Label20 
         Caption         =   "Min Fader"
         Height          =   255
         Left            =   2550
         TabIndex        =   53
         ToolTipText     =   "Minimize Fader"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "Refresh Colors"
         Height          =   255
         Left            =   2860
         TabIndex        =   52
         ToolTipText     =   "Refresh Colors"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1a 
         Caption         =   "R"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   300
         Width           =   135
      End
      Begin VB.Label Label1a 
         Caption         =   "B"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   1500
         Width           =   135
      End
      Begin VB.Label Label1a 
         Caption         =   "G"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   900
         Width           =   135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copy Text"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Copy Text"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Add Color"
         Height          =   255
         Left            =   3060
         TabIndex        =   17
         ToolTipText     =   "Add Color"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Remove Color"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         ToolTipText     =   "Remove Color"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Replace Color"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         ToolTipText     =   "Replace Color"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Insert Color"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         ToolTipText     =   "Insert Color"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clear Text"
         Height          =   255
         Left            =   940
         TabIndex        =   12
         ToolTipText     =   "Clear Text"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fade"
         Height          =   255
         Left            =   3220
         TabIndex        =   11
         Tag             =   "0"
         ToolTipText     =   "Fade"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Advertise"
         Height          =   255
         Left            =   3045
         TabIndex        =   10
         Tag             =   "0"
         ToolTipText     =   "Advertise"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Fader"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         ToolTipText     =   "Exit Fader"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Webpage"
         Height          =   255
         Left            =   1740
         TabIndex        =   8
         ToolTipText     =   "Coolest Fader Homepage!"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   1440
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2415
      Index           =   2
      Left            =   420
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Frame AFrame1 
         Caption         =   "Type of Coding"
         Height          =   855
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton AOption1 
            Caption         =   "HTML"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "HTML"
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton AOption2 
            Caption         =   "ANSI"
            Height          =   255
            Left            =   1440
            TabIndex        =   31
            ToolTipText     =   "ANSI"
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton AOption3 
            Caption         =   "Yahoo! Chat Color Codes"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Yahoo! Chat Color Codes"
            Top             =   240
            Value           =   -1  'True
            Width           =   2145
         End
      End
      Begin VB.Frame AFrame2 
         Caption         =   "Other"
         Height          =   855
         Left            =   1440
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
         Begin VB.CheckBox ACheck1 
            Caption         =   "Clear on Copy"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            ToolTipText     =   "Clear on Copy"
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox ACheck2 
            Caption         =   "Copy on Fade"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Copy on Fade"
            Top             =   480
            Width           =   1300
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extras"
         Height          =   2055
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
         Begin VB.CheckBox Check7 
            Caption         =   "Screwed"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1680
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bold"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Bold"
            Top             =   480
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Italics"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Italiacs"
            Top             =   720
            Width           =   795
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Underline"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Underline"
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox Check4 
            Caption         =   "On Top"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Always On Top"
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Reverse"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Reverse"
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Alt Caps"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            ToolTipText     =   "Alternating Caps"
            Top             =   1440
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   2775
      Index           =   4
      Left            =   240
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Text            =   "Frames.frx":0446
         ToolTipText     =   "Version Info"
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   4185
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Changes in the Versions:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"Frames.frx":0E17
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   2775
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   855
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "Preview"
         Top             =   1800
         Width           =   3975
         ExtentX         =   7011
         ExtentY         =   1508
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   -1  'True
         NoClientEdge    =   -1  'True
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   1800
         Width           =   3975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   195
         Left            =   3960
         TabIndex        =   55
         Top             =   2460
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   37
         ToolTipText     =   "Example"
         Top             =   2640
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H000000FF&
         Height          =   1485
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         ToolTipText     =   "Source"
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Credits"
      Height          =   2775
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Credits to: Borntopumpgas, CertWiz and Wargod!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Credits"
         Top             =   360
         Width           =   3975
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3255
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fade!"
            Object.ToolTipText     =   "Fade!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Object.ToolTipText     =   "Preview"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Object.ToolTipText     =   "Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Credits"
            Object.ToolTipText     =   "Credits"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   6240
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   375
      Left            =   6360
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu FadeMenu 
      Caption         =   "Fade"
      Visible         =   0   'False
      Begin VB.Menu FadeColors2 
         Caption         =   "Fade!"
      End
      Begin VB.Menu AddColor2 
         Caption         =   "Add Color"
      End
      Begin VB.Menu InsertColor2 
         Caption         =   "Insert Color"
      End
      Begin VB.Menu RemoveColor2 
         Caption         =   "Remove Color"
      End
      Begin VB.Menu ReplaceColor2 
         Caption         =   "Replace Color"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim li As Long
Dim r(256) As Double
Dim g(256) As Double
Dim b(256) As Double
Dim n As Long
Dim s(52) As String
Dim H2T(255) As String

Private Function FileExist(ByRef sFname) As Boolean

        If Len(Dir(sFname, 16)) Then FileExist = True Else FileExist = False

End Function

Public Sub PowerStart()
Directory$ = App.Path & "\preview.dat"
If Not FileExist(Directory$) Then
Exit Sub
Else
    Dim MyString As String
       On Error Resume Next
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
           Text4.Text = MyString$
           Wend
           Close #1
End If
End Sub

Public Sub PowerUp()
If Len(Text4.Text) > 0 Then
Dim Directory As String
If List1.ListCount > 1 Then
Directory$ = App.Path & "\preview.htm"
Dim s As String
If Check1.Value = 1 Then
    s$ = s$ + "<b>"
  End If
  If Check2.Value = 1 Then
    s$ = s$ + "<i>"
  End If
  If Check3.Value = 1 Then
    s$ = s$ + "<u>"
  End If
  Dim ColorR As Double
  Dim ColorG As Double
  Dim ColorB As Double
  Dim NumColors As Double
  Dim NumFades As Double
  Dim NumLetters As Double
  Dim CurrFade As Integer
  Dim CurrLetter As String
  Dim LettersPerFade As Double
  Dim FadeLetter As Double
  Dim RDelta As Double
  Dim GDelta As Double
  Dim BDelta As Double
  NumColors = List1.ListCount
  NumFades = NumColors - 1
  NumLetters = Len(Text4.Text)
  ColorR = r(0)
  ColorG = g(0)
  ColorB = b(0)
  LettersPerFade = NumLetters / NumFades
  RDelta = (r(1) - r(0)) / LettersPerFade
  GDelta = (g(1) - g(0)) / LettersPerFade
  BDelta = (b(1) - b(0)) / LettersPerFade
  FadeLetter = 0
  CurrFade = 0
  Picture1.Cls
  s$ = ""
  For Letter = 1 To NumLetters
    CurrLetter = Mid$(Text4.Text, Letter, 1)
    If ColorR > 255 Then ColorR = 255
    If ColorR < 0 Then ColorR = 0
    If ColorG > 255 Then ColorG = 255
    If ColorG < 0 Then ColorG = 0
    If ColorB > 255 Then ColorB = 255
    If ColorB < 0 Then ColorB = 0
    Picture1.ForeColor = RGB(Int(ColorR), Int(ColorG), Int(ColorB))
    s$ = s$ & "<FONT COLOR=" & H2T(Int(ColorR)) & H2T(Int(ColorG)) & H2T(Int(ColorB)) & ">" & CurrLetter & "</FONT>"
    ColorR = ColorR + RDelta
    ColorG = ColorG + GDelta
    ColorB = ColorB + BDelta
    Picture1.Print CurrLetter;
    If FadeLetter >= LettersPerFade Then
      FadeLetter = FadeLetter - LettersPerFade
      CurrFade = CurrFade + 1
      RDelta = (r(CurrFade + 1) - r(CurrFade)) / LettersPerFade
      GDelta = (g(CurrFade + 1) - g(CurrFade)) / LettersPerFade
      BDelta = (b(CurrFade + 1) - b(CurrFade)) / LettersPerFade
      ColorR = r(CurrFade)
      ColorG = g(CurrFade)
      ColorB = b(CurrFade)
    Else
      FadeLetter = FadeLetter + 1
    End If
  Next Letter
If Not FileExist(Directory$) Then
Exit Sub
Else
       On Error Resume Next
       Open Directory$ For Output As #1
           Print #1, s$
       Close #1
End If
End If
End If
Directory$ = App.Path & "\preview.dat"
If Not FileExist(Directory$) Then
Exit Sub
Else
       On Error Resume Next
       Open Directory$ For Output As #1
           Print #1, Text4.Text
       Close #1
End If
End Sub

Public Sub Sav()
Directory$ = App.Path & "\fader.dat"
       Dim SaveList As Integer
       On Error Resume Next
       Open Directory$ For Output As #1
       For SaveList = 0 To List1.ListCount - 1
           Print #1, r(SaveList)
           Print #1, g(SaveList)
           Print #1, b(SaveList)
       Next SaveList
       Close #1
End Sub

Public Sub Opn()
Directory$ = App.Path & "\fader.dat"
If Not FileExist(Directory$) Then
Exit Sub
Else
    Dim MyString As String
    Dim a As Long
    Dim c As Long
    Dim i As Integer
    a = 1
    c = 0
       On Error Resume Next
       Open Directory$ For Input As #1
       While Not EOF(1)
           Input #1, MyString$
           DoEvents
               If a = 1 Then
               r(c) = MyString$
               a = 2
               ElseIf a = 2 Then
               g(c) = MyString$
               a = 3
               ElseIf a = 3 Then
               b(c) = MyString$
               a = 1
               c = c + 1
               End If
           Wend
           Close #1
For i = 0 To c - 1
List1.AddItem Str$(r(i)) + "," + Str$(g(i)) + "," + Str$(b(i))
List2.AddItem Str$(r(i)) + "," + Str$(g(i)) + "," + Str$(b(i))
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

Public Sub a()
li = Form2.List1.ListIndex
s(li) = Form2.Text2.Text
End Sub

Public Sub Elite()
Text4.Text = sReplaceCharacters(Text4.Text, "a", s(1))
Text4.Text = sReplaceCharacters(Text4.Text, "b", s(2))
Text4.Text = sReplaceCharacters(Text4.Text, "c", s(3))
Text4.Text = sReplaceCharacters(Text4.Text, "d", s(4))
Text4.Text = sReplaceCharacters(Text4.Text, "e", s(5))
Text4.Text = sReplaceCharacters(Text4.Text, "f", s(6))
Text4.Text = sReplaceCharacters(Text4.Text, "g", s(7))
Text4.Text = sReplaceCharacters(Text4.Text, "h", s(8))
Text4.Text = sReplaceCharacters(Text4.Text, "i", s(9))
Text4.Text = sReplaceCharacters(Text4.Text, "j", s(10))
Text4.Text = sReplaceCharacters(Text4.Text, "k", s(11))
Text4.Text = sReplaceCharacters(Text4.Text, "l", s(12))
Text4.Text = sReplaceCharacters(Text4.Text, "m", s(13))
Text4.Text = sReplaceCharacters(Text4.Text, "n", s(14))
Text4.Text = sReplaceCharacters(Text4.Text, "o", s(15))
Text4.Text = sReplaceCharacters(Text4.Text, "p", s(16))
Text4.Text = sReplaceCharacters(Text4.Text, "q", s(17))
Text4.Text = sReplaceCharacters(Text4.Text, "r", s(18))
Text4.Text = sReplaceCharacters(Text4.Text, "s", s(19))
Text4.Text = sReplaceCharacters(Text4.Text, "t", s(20))
Text4.Text = sReplaceCharacters(Text4.Text, "u", s(21))
Text4.Text = sReplaceCharacters(Text4.Text, "v", s(22))
Text4.Text = sReplaceCharacters(Text4.Text, "w", s(23))
Text4.Text = sReplaceCharacters(Text4.Text, "x", s(24))
Text4.Text = sReplaceCharacters(Text4.Text, "y", s(25))
Text4.Text = sReplaceCharacters(Text4.Text, "z", s(26))
Text4.Text = sReplaceCharacters(Text4.Text, "A", s(27))
Text4.Text = sReplaceCharacters(Text4.Text, "B", s(28))
Text4.Text = sReplaceCharacters(Text4.Text, "C", s(29))
Text4.Text = sReplaceCharacters(Text4.Text, "D", s(30))
Text4.Text = sReplaceCharacters(Text4.Text, "E", s(31))
Text4.Text = sReplaceCharacters(Text4.Text, "F", s(32))
Text4.Text = sReplaceCharacters(Text4.Text, "G", s(33))
Text4.Text = sReplaceCharacters(Text4.Text, "H", s(34))
Text4.Text = sReplaceCharacters(Text4.Text, "I", s(35))
Text4.Text = sReplaceCharacters(Text4.Text, "J", s(36))
Text4.Text = sReplaceCharacters(Text4.Text, "K", s(37))
Text4.Text = sReplaceCharacters(Text4.Text, "L", s(38))
Text4.Text = sReplaceCharacters(Text4.Text, "M", s(39))
Text4.Text = sReplaceCharacters(Text4.Text, "N", s(40))
Text4.Text = sReplaceCharacters(Text4.Text, "O", s(41))
Text4.Text = sReplaceCharacters(Text4.Text, "P", s(42))
Text4.Text = sReplaceCharacters(Text4.Text, "Q", s(43))
Text4.Text = sReplaceCharacters(Text4.Text, "R", s(44))
Text4.Text = sReplaceCharacters(Text4.Text, "S", s(45))
Text4.Text = sReplaceCharacters(Text4.Text, "T", s(46))
Text4.Text = sReplaceCharacters(Text4.Text, "U", s(47))
Text4.Text = sReplaceCharacters(Text4.Text, "V", s(48))
Text4.Text = sReplaceCharacters(Text4.Text, "W", s(49))
Text4.Text = sReplaceCharacters(Text4.Text, "X", s(50))
Text4.Text = sReplaceCharacters(Text4.Text, "Y", s(51))
Text4.Text = sReplaceCharacters(Text4.Text, "Z", s(52))
Text4.Text = sReplaceCharacters(Text4.Text, "ae", s(0))
End Sub

Public Sub Caps()
Dim i As Integer
Dim s As String
s = ""
For i = 1 To Len(Text4.Text)
  keyval = Asc(Mid$(Text4.Text, i, 1))
  If (keyval >= 96 And keyval < 96 + 26) Or (keyval >= 64 And keyval < 64 + 26) Then
    If (i And 1) = 1 Then
      If keyval < 96 Then
        s = s + Chr$(96 + keyval - 64)
      Else
        s = s + Chr$(keyval)
      End If
    Else
      If keyval >= 96 Then
        s = s + Chr$(64 + keyval - 96)
      Else
        s = s + Chr$(keyval)
      End If
    End If
  Else
    s = s + Chr$(keyval)
  End If
Next i
Text4.Text = s
End Sub

Public Sub Preview()
If WebBrowser1.Tag = "1" Then
WebBrowser1.Visible = True
Text2.Visible = False
WebBrowser1.Navigate App.Path & "\preview.htm"
Else
Text2.Visible = True
WebBrowser1.Visible = False
End If
End Sub

Private Sub AddColor2_Click()
Label4_Click
End Sub

Private Sub AOption1_Click()
Text4.MaxLength = 0
End Sub

Private Sub AOption2_Click()
Text4.MaxLength = 0
End Sub

Private Sub AOption3_Click()
If AOption3.Value = True Then
Text4.MaxLength = 45
Else
Text4.MaxLength = 0
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
Form1.Hide
Form2.Show
End If
End Sub

Private Sub Check7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check7.ForeColor = &HFF00&
Check5.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Check6.ForeColor = &H0&
End Sub

Private Sub Command2_Click()
Preview
End Sub

Private Sub FadeColors2_Click()
Label12_Click
End Sub

Private Sub Form_Activate()
Timer1.Enabled = True
End Sub

Private Sub Form_GotFocus()
Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
Timer1.Enabled = True
End Sub

Private Sub HScroll1_Change()
Dim X%, x2!, Y%, i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture2.ScaleWidth / 100
   Y% = Picture2.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = 0
   blue1% = Str$(HScroll3.Value)
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = 255
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture2.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next

   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture3.ScaleWidth / 100
   Y% = Picture3.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = 0
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = 255
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture3.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next
  Label1.Caption = Str$(HScroll1.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Dim X%, x2!, Y%, i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture1a.ScaleWidth / 100
   Y% = Picture1a.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = 0 ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = Str$(HScroll3.Value)
   red2% = 255 ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture1a.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next

   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture3.ScaleWidth / 100
   Y% = Picture3.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = 0
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = 255
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture3.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next
  Label2.Caption = Str$(HScroll2.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Dim X%, x2!, Y%, i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture2.ScaleWidth / 100
   Y% = Picture2.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = 0
   blue1% = Str$(HScroll3.Value)
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = 255
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture2.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next

   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture1a.ScaleWidth / 100
   Y% = Picture1a.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = 0 ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = Str$(HScroll3.Value)
   red2% = 255 ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture1a.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next
  Label3.Caption = Str$(HScroll3.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub


Private Sub InsertColor2_Click()
Label8_Click
End Sub

Private Sub Label12_Click()
If List1.ListCount < 2 Then
  temp = MsgBox("You must have atleast 2 colors for the fade to work!", vbExclamation Or vbOKOnly, "Error")
  Exit Sub
End If
If Text4.Text = "" Then
  temp = MsgBox("You have to have text to fade!", vbExclamation Or vbOKOnly, "Error")
  Exit Sub
End If
  If Check1.Value = 1 Then
    Text5.Text = Text5.Text + "<b>"
  Else
    Text5.Text = Text5.Text
  End If
  If Check2.Value = 1 Then
    Text5.Text = Text5.Text + "<i>"
  Else
    Text5.Text = Text5.Text
  End If
  If Check3.Value = 1 Then
    Text5.Text = Text5.Text + "<u>"
  Else
    Text5.Text = Text5.Text
  End If
  If Check5.Value = 1 Then
    Command1_Click
  End If
  If Check6.Value = 1 Then
    Caps
  End If
  If Check7.Value = 1 Then
    Elite
  End If
    If Form1.AOption3.Value = True Then
  Dim ColorR As Double
  Dim ColorG As Double
  Dim ColorB As Double
  Dim NumColors As Double
  Dim NumFades As Double
  Dim NumLetters As Double
  Dim CurrFade As Integer
  Dim CurrLetter As String
  Dim LettersPerFade As Double
  Dim FadeLetter As Double
  Dim RDelta As Double
  Dim GDelta As Double
  Dim BDelta As Double
  NumColors = List1.ListCount
  NumFades = NumColors - 1
  NumLetters = Len(Text4.Text)
  ColorR = r(0)
  ColorG = g(0)
  ColorB = b(0)
  LettersPerFade = NumLetters / NumFades
  RDelta = (r(1) - r(0)) / LettersPerFade
  GDelta = (g(1) - g(0)) / LettersPerFade
  BDelta = (b(1) - b(0)) / LettersPerFade
  FadeLetter = 0
  CurrFade = 0
  Picture1.Cls
  Text5.Text = ""
  For Letter = 1 To NumLetters
    CurrLetter = Mid$(Text4.Text, Letter, 1)
    If ColorR > 255 Then ColorR = 255
    If ColorR < 0 Then ColorR = 0
    If ColorG > 255 Then ColorG = 255
    If ColorG < 0 Then ColorG = 0
    If ColorB > 255 Then ColorB = 255
    If ColorB < 0 Then ColorB = 0
    Picture1.ForeColor = RGB(Int(ColorR), Int(ColorG), Int(ColorB))
    Text5.Text = Text5.Text & "<#" & H2T(Int(ColorR)) & H2T(Int(ColorG)) & H2T(Int(ColorB)) & ">" & CurrLetter
    ColorR = ColorR + RDelta
    ColorG = ColorG + GDelta
    ColorB = ColorB + BDelta
    Picture1.Print CurrLetter;
    If FadeLetter >= LettersPerFade Then
      FadeLetter = FadeLetter - LettersPerFade
      CurrFade = CurrFade + 1
      RDelta = (r(CurrFade + 1) - r(CurrFade)) / LettersPerFade
      GDelta = (g(CurrFade + 1) - g(CurrFade)) / LettersPerFade
      BDelta = (b(CurrFade + 1) - b(CurrFade)) / LettersPerFade
      ColorR = r(CurrFade)
      ColorG = g(CurrFade)
      ColorB = b(CurrFade)
    Else
      FadeLetter = FadeLetter + 1
    End If
  Next Letter
  ElseIf Form1.AOption2.Value = True Then
  If List1.ListCount < 2 Then
  temp = MsgBox("You must have at least 2 colors in your list to fade!", vbExclamation Or vbOKOnly, "Warning!!!")
  Exit Sub
End If
If Text4.Text = "" Then
  temp = MsgBox("You must type a message to fade!", vbExclamation Or vbOKOnly, "Warning!!!")
  Exit Sub
End If
  NumColors = List1.ListCount
  NumFades = NumColors - 1
  NumLetters = Len(Text4.Text)
  ColorR = r(0)
  ColorG = g(0)
  ColorB = b(0)
  LettersPerFade = NumLetters / NumFades
  RDelta = (r(1) - r(0)) / LettersPerFade
  GDelta = (g(1) - g(0)) / LettersPerFade
  BDelta = (b(1) - b(0)) / LettersPerFade
  FadeLetter = 0
  CurrFade = 0
  Picture1.Cls
  Text5.Text = ""
  For Letter = 1 To NumLetters
    CurrLetter = Mid$(Text4.Text, Letter, 1)
    If ColorR > 255 Then ColorR = 255
    If ColorR < 0 Then ColorR = 0
    If ColorG > 255 Then ColorG = 255
    If ColorG < 0 Then ColorG = 0
    If ColorB > 255 Then ColorB = 255
    If ColorB < 0 Then ColorB = 0
    Picture1.ForeColor = RGB(Int(ColorR), Int(ColorG), Int(ColorB))
    Text5.Text = Text5.Text & "[#" & H2T(Int(ColorR)) & H2T(Int(ColorG)) & H2T(Int(ColorB)) & CurrLetter
    ColorR = ColorR + RDelta
    ColorG = ColorG + GDelta
    ColorB = ColorB + BDelta
    Picture1.Print CurrLetter;
    If FadeLetter >= LettersPerFade Then
      FadeLetter = FadeLetter - LettersPerFade
      CurrFade = CurrFade + 1
      RDelta = (r(CurrFade + 1) - r(CurrFade)) / LettersPerFade
      GDelta = (g(CurrFade + 1) - g(CurrFade)) / LettersPerFade
      BDelta = (b(CurrFade + 1) - b(CurrFade)) / LettersPerFade
      ColorR = r(CurrFade)
      ColorG = g(CurrFade)
      ColorB = b(CurrFade)
    Else
      FadeLetter = FadeLetter + 1
    End If
  Next Letter
  Else
  If List1.ListCount < 2 Then
  temp = MsgBox("You must have at least 2 colors in your list to fade!", vbExclamation Or vbOKOnly, "Warning!!!")
  Exit Sub
End If
If Text4.Text = "" Then
  temp = MsgBox("You must type a message to fade!", vbExclamation Or vbOKOnly, "Warning!!!")
  Exit Sub
End If
  NumColors = List1.ListCount
  NumFades = NumColors - 1
  NumLetters = Len(Text4.Text)
  ColorR = r(0)
  ColorG = g(0)
  ColorB = b(0)
  LettersPerFade = NumLetters / NumFades
  RDelta = (r(1) - r(0)) / LettersPerFade
  GDelta = (g(1) - g(0)) / LettersPerFade
  BDelta = (b(1) - b(0)) / LettersPerFade
  FadeLetter = 0
  CurrFade = 0
  Picture1.Cls
  Text5.Text = ""
  For Letter = 1 To NumLetters
    CurrLetter = Mid$(Text4.Text, Letter, 1)
    If ColorR > 255 Then ColorR = 255
    If ColorR < 0 Then ColorR = 0
    If ColorG > 255 Then ColorG = 255
    If ColorG < 0 Then ColorG = 0
    If ColorB > 255 Then ColorB = 255
    If ColorB < 0 Then ColorB = 0
    Picture1.ForeColor = RGB(Int(ColorR), Int(ColorG), Int(ColorB))
    Text5.Text = Text5.Text & "<FONT COLOR=" & H2T(Int(ColorR)) & H2T(Int(ColorG)) & H2T(Int(ColorB)) & ">" & CurrLetter & "</FONT>"
    ColorR = ColorR + RDelta
    ColorG = ColorG + GDelta
    ColorB = ColorB + BDelta
    Picture1.Print CurrLetter;
    If FadeLetter >= LettersPerFade Then
      FadeLetter = FadeLetter - LettersPerFade
      CurrFade = CurrFade + 1
      RDelta = (r(CurrFade + 1) - r(CurrFade)) / LettersPerFade
      GDelta = (g(CurrFade + 1) - g(CurrFade)) / LettersPerFade
      BDelta = (b(CurrFade + 1) - b(CurrFade)) / LettersPerFade
      ColorR = r(CurrFade)
      ColorG = g(CurrFade)
      ColorB = b(CurrFade)
    Else
      FadeLetter = FadeLetter + 1
    End If
  Next Letter
End If
If Label13.Tag = "1" Then
Label12.Tag = Label12.Tag + 1
ElseIf Label13.Tag = "0" Then
Label12.Tag = Label13.Tag
End If
If Form1.ACheck2.Value = 1 Then
  If Label12.Tag = "1" Then
  Clipboard.SetText Text5.Text + "  http://fly.to/coolestfader/"
  Else
  Clipboard.SetText Text5.Text
  End If
    If Form1.ACheck1.Value = 1 Then
    Text4.Text = ""
    Text5.Text = ""
    End If
End If
WebBrowser1.Tag = 1
PowerUp
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = &HFF00FF
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label4.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label13_Click()
Dim i As Integer
For i = 0 To List1.ListCount - 1
r(i) = ""
g(i) = ""
b(i) = ""
Next i
List1.Clear
HScroll1.Value = 255
HScroll2.Value = 0
HScroll3.Value = 0
Label4_Click
HScroll1.Value = 0
HScroll2.Value = 0
HScroll3.Value = 255
Label4_Click
HScroll1.Value = 0
HScroll2.Value = 255
HScroll3.Value = 0
Label4_Click
Text4.Text = "I am using " + Form1.Caption + " by Coolestmon!"
Label13.Tag = Label13.Tag + 1
Label12_Click
Label13.Tag = Label13.Tag - 1
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label4.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label14_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo, "Exit") = vbNo Then
Exit Sub
Else
Sav
End
End If
End Sub

Private Sub Label19_Click()
Dim X%, x2!, Y%, i%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture1a.ScaleWidth / 100
   Y% = Picture1a.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = 0 ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = Str$(HScroll3.Value)
   red2% = 255 ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture1a.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next
   
   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture2.ScaleWidth / 100
   Y% = Picture2.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = 0
   blue1% = Str$(HScroll3.Value)
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = 255
   blue2% = Str$(HScroll3.Value)
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture2.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next

   X% = 0
   ' find the length of the picturebox and cut it into 100 pieces
   x2 = Picture3.ScaleWidth / 100
   Y% = Picture3.ScaleHeight
   ' setting how much red, green, and blue goes into each of the two
   '     colors
   red1% = Str$(HScroll1.Value) ' the amount of red in color one
   green1% = Str$(HScroll2.Value)
   blue1% = 0
   red2% = Str$(HScroll1.Value) ' the amount of red in color two
   green2% = Str$(HScroll2.Value)
   blue2% = 255
   ' cut the difference between the two colors into 100 pieces
   pat1 = (red2% - red1%) / 100
   pat2 = (green2% - green1%) / 100
   pat3 = (blue2% - blue1%) / 100
   ' set the c variables at the starting colors
   c1 = red1%
   c2 = green1%
   c3 = blue1%
   ' draw 100 different lines on the picturebox


   For i% = 1 To 100
       Picture3.Line (X%, 0)-(X% + x2, Y%), RGB(c1, c2, c3), BF
       X% = X% + x2 ' draw the next line one step up from the old step
       c1 = c1 + pat1 ' make the c variable equal 2 it's next step
       c2 = c2 + pat2
       c3 = c3 + pat3
   Next
Command2_Click
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = &HFFFF&
Label4.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label5.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label20_Click()
Form1.WindowState = 1
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = &HFF0000
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label4.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label14.ForeColor = &H0&
End Sub


Private Sub Label34_Click()
End Sub

Private Sub Label28_Click()
End Sub

Private Sub Label23_Click()
End Sub

Private Sub List1_Click()
If List1.ListIndex < 0 Then
Exit Sub
Else
List2.ListIndex = List1.ListIndex
End If
If List1.ListIndex < 0 Then
Exit Sub
Else
li = List1.ListIndex
HScroll1.Value = r(li)
HScroll2.Value = g(li)
HScroll3.Value = b(li)
End If
End Sub

Private Sub List1_ItemCheck(Item As Integer)
List1_Click
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If List1.ListIndex < 1 Then
Exit Sub
Else
List2.ListIndex = List1.ListIndex
End If
If List1.ListIndex < 0 Then
Exit Sub
Else
li = List1.ListIndex
HScroll1.Value = r(li)
HScroll2.Value = g(li)
HScroll3.Value = b(li)
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If List1.ListIndex < 1 Then
Exit Sub
Else
List2.ListIndex = List1.ListIndex
End If
If List1.ListIndex < 0 Then
Exit Sub
Else
li = List1.ListIndex
HScroll1.Value = r(li)
HScroll2.Value = g(li)
HScroll3.Value = b(li)
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu FadeMenu
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List1.ListIndex < 1 Then
Exit Sub
Else
List2.ListIndex = List1.ListIndex
End If
End Sub

Private Sub List5_Click()

End Sub

Private Sub Picture1a_Click()
HScroll1.SetFocus
End Sub

Private Sub Picture2_Click()
HScroll2.SetFocus
End Sub

Private Sub Picture3_Click()
HScroll3.SetFocus
End Sub


Private Sub RemoveColor2_Click()
Label5_Click
End Sub

Private Sub ReplaceColor2_Click()
Label6_Click
End Sub

Private Sub Timer1_Timer()
Label19_Click
Timer1.Enabled = False
End Sub


Private Sub AboutT_Click()
Form1.TabStrip1.SelectedItem.Index = 5
End Sub

Private Sub AFrame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub AFrame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub AOption1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AOption1.ForeColor = &HFFFF00
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub AOption2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AOption2.ForeColor = &HFFFF00
AOption3.ForeColor = &H0&
AOption1.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub AOption3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AOption3.ForeColor = &HFFFF00
AOption2.ForeColor = &H0&
AOption1.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub


Private Sub ACheck1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &HFF00&
ACheck2.ForeColor = &H0&
End Sub

Private Sub ACheck2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck2.ForeColor = &HFF00&
ACheck1.ForeColor = &H0&
End Sub
Private Sub Check1_Click()
  Picture1.Font.Bold = Check1.Value
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &HFF00&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub Check2_Click()
  Picture1.Font.Italic = Check2.Value
End Sub


Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check2.ForeColor = &HFF00&
Check1.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub Check3_Click()
  Picture1.Font.Underline = Check3.Value
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check3.ForeColor = &HFF00&
Check2.ForeColor = &H0&
Check1.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub Check4_Click()
If Check4.Value = 0 Then
Notontop Me
Else
Ontop Me
End If
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check4.ForeColor = &HFF00&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check1.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub



Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &HFF00&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
End Sub

Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check6.ForeColor = &HFF00&
Check5.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Check7.ForeColor = &H0&
End Sub



Private Sub Command1_Click()
strin$ = Text4.Text
For i% = 1 To Len(strin$)
stringy$ = Mid$(strin$, i%, 1)
final$ = stringy$ + final$
Next i%
Text4.Text = final$
End Sub

Private Sub CreditsT_Click()
TabStrip1.SelectedItem.Index = 4
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Fade_Click()
Label12_Click
End Sub

Private Sub FadeT_Click()
TabStrip1.SelectedItem.Index = 1
End Sub

Private Sub Form_Load()
  Dim TempH2T(15) As String
  Dim HexIndex As Integer
  TempH2T$(0) = "0"
  TempH2T$(1) = "1"
  TempH2T$(2) = "2"
  TempH2T$(3) = "3"
  TempH2T$(4) = "4"
  TempH2T$(5) = "5"
  TempH2T$(6) = "6"
  TempH2T$(7) = "7"
  TempH2T$(8) = "8"
  TempH2T$(9) = "9"
  TempH2T$(10) = "A"
  TempH2T$(11) = "B"
  TempH2T$(12) = "C"
  TempH2T$(13) = "D"
  TempH2T$(14) = "E"
  TempH2T$(15) = "F"
  For HexIndex = 0 To 255
    H2T(HexIndex) = TempH2T(Int(HexIndex / 16)) & TempH2T(HexIndex And 15)
  Next HexIndex
PowerStart
Opn
If Len(Text4.Text) > 0 Then
  If List1.ListCount > 1 Then
  Label12_Click
  End If
End If
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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Form_Terminate()
Sav
PowerUp
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sav
PowerUp
End
End Sub
Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
Check7.ForeColor = &H0&
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
Check7.ForeColor = &H0&
End Sub

Private Sub HScroll1_Scroll()
  Label1.Caption = Str$(HScroll1.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
  HScroll1_Change
End Sub

Private Sub HScroll2_Scroll()
  Label2.Caption = Str$(HScroll2.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
  HScroll2_Change
End Sub

Private Sub HScroll3_Scroll()
  Label3.Caption = Str$(HScroll3.Value)
  Shape2.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
  HScroll3_Change
End Sub

Private Sub Label10_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label4.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label11_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label4.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = &HFF0000
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label4.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label15_Click()
Ret = Shell("Start.exe " & "http://www.fly.to/coolestfader/", 0)
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &HFFFF&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label4_Click()
If List1.ListCount > 255 Then
MsgBox "Too many colors!", 48, "Error"
Exit Sub
Else
 List2.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value)
  List2.ListIndex = List2.ListCount - 1
  li = List2.ListCount - 1
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
 List1.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value)
  List1.ListIndex = List1.ListCount - 1
  li = List1.ListCount - 1
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label5_Click()
  li = List2.ListIndex
  If List2.ListIndex > List2.ListCount - 1 Or List2.ListIndex < 0 Then Exit Sub
  List2.RemoveItem List2.ListIndex
  For colorindex = li To List2.ListCount - 1
    r(colorindex) = r(colorindex + 1)
    g(colorindex) = g(colorindex + 1)
    b(colorindex) = b(colorindex + 1)
  Next colorindex
  If li = List2.ListCount Then li = li - 1
  List2.ListIndex = li
  li = List1.ListIndex
  If List1.ListIndex > List1.ListCount - 1 Or List1.ListIndex < 0 Then Exit Sub
  List1.RemoveItem List1.ListIndex
  For colorindex = li To List1.ListCount - 1
    r(colorindex) = r(colorindex + 1)
    g(colorindex) = g(colorindex + 1)
    b(colorindex) = b(colorindex + 1)
  Next colorindex
  If li = List1.ListCount Then li = li - 1
  List1.ListIndex = li
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF&
Label4.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label6_Click()
  li = List2.ListIndex
  If List2.ListIndex > List2.ListCount - 1 Or List2.ListIndex < 0 Then Exit Sub
  List2.RemoveItem li
  List2.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value), li
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
  List2.ListIndex = li
  li = List1.ListIndex
  If List1.ListIndex > List1.ListCount - 1 Or List1.ListIndex < 0 Then Exit Sub
  List1.RemoveItem li
  List1.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value), li
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
  List1.ListIndex = li
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label4.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label7_Click()
If Label12.Tag = "1" Then
Clipboard.SetText Text5.Text + "  http://fly.to/coolestfader/"
Else
Clipboard.SetText Text5.Text
End If
  If Form1.ACheck1.Value = 1 Then
  Text4.Text = ""
  Text5.Text = ""
  End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label4.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label8_Click()
If List1.ListCount > 255 Then
MsgBox "Too many colors!", 48, "Error"
Exit Sub
 If List2.ListIndex > List2.ListCount - 1 Or List2.ListIndex < 0 Then
    Label4_Click
    Exit Sub
  End If
  List2.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value), li
  For colorindex = List2.ListCount + 1 To li + 1 Step -1
    r(colorindex) = r(colorindex - 1)
    g(colorindex) = g(colorindex - 1)
    b(colorindex) = b(colorindex - 1)
  Next colorindex
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
  If li = List2.ListCount Then li = li - 1
  List2.ListIndex = li
  If List1.ListIndex > List1.ListCount - 1 Or List1.ListIndex < 0 Then
    Label4_Click
    Exit Sub
  End If
  List1.AddItem Str$(HScroll1.Value) + "," + Str$(HScroll2.Value) + "," + Str$(HScroll3.Value), li
  For colorindex = List1.ListCount + 1 To li + 1 Step -1
    r(colorindex) = r(colorindex - 1)
    g(colorindex) = g(colorindex - 1)
    b(colorindex) = b(colorindex - 1)
  Next colorindex
  r(li) = HScroll1.Value
  g(li) = HScroll2.Value
  b(li) = HScroll3.Value
  If li = List1.ListCount Then li = li - 1
  List1.ListIndex = li
End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label4.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub Label9_Click()
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFF&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label4.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
Label19.ForeColor = &H0&
Label20.ForeColor = &H0&
End Sub

Private Sub OptionsMenu_Click()

End Sub

Private Sub PreviewT_Click()
TabStrip1.SelectedItem.Index = 2
End Sub

Private Sub TabStrip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H0&
Label5.ForeColor = &H0&
Label6.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H0&
Label9.ForeColor = &H0&
Label10.ForeColor = &H0&
Label11.ForeColor = &H0&
Label12.ForeColor = &H0&
Label13.ForeColor = &H0&
Label14.ForeColor = &H0&
Check1.ForeColor = &H0&
Check2.ForeColor = &H0&
Check3.ForeColor = &H0&
Check4.ForeColor = &H0&
Label15.ForeColor = &H0&
Check5.ForeColor = &H0&
Check6.ForeColor = &H0&
AOption1.ForeColor = &H0&
AOption2.ForeColor = &H0&
AOption3.ForeColor = &H0&
ACheck1.ForeColor = &H0&
ACheck2.ForeColor = &H0&
Label20.ForeColor = &H0&
Timer1.Enabled = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub
Private Sub TabStrip1_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To TabStrip1.Tabs.Count - 1
        If i = TabStrip1.SelectedItem.Index - 1 Then
        Frame1(i).Visible = True
        Else
        Frame1(i).Visible = False
        End If
        Next
End Sub


Private Sub Timer2_Timer()
MsgBox "Here goes nothin!", 48, "ok..."
Text5.Text = Text5.Text + "  http://fly.to/coolestfader"
Timer2.Enabled = False
End Sub
