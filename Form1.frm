VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4635
   ScaleWidth      =   5640
   Begin RichTextLib.RichTextBox rchMainText 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "www.Bizzit.com"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2175
      TabIndex        =   2
      Top             =   4200
      Width           =   1170
   End
   Begin VB.Label lblCurrentWord 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   30
      TabIndex        =   1
      Top             =   3285
      Width           =   5520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' Return the word the mouse is over.
Public Function RichWordOver(rch As RichTextBox, X As Single, y As Single) As String
On Error Resume Next
    Dim pt        As POINTAPI
    Dim pos       As Long
    Dim start_pos As Long
    Dim end_pos   As Long
    Dim ch        As String
    Dim txt       As String
    Dim txtlen    As Long

    ' Convert the position to pixels.
    pt.X = X \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hwnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If ch = " " Or ch = Chr(13) Then Exit For
    Next start_pos
    
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If ch = " " Or ch = Chr(13) Then Exit For
    Next end_pos
    
    end_pos = end_pos - 1

    If start_pos <= end_pos Then RichWordOver = Mid$(txt, start_pos, end_pos - start_pos + 1)
End Function



Private Sub Form_Load()
    rchMainText.Text = _
        "Goto www.Bizzit.com" & vbCrLf & vbCrLf & _
        "We have spend 9 months making a very nice program for you so programmers can come in real easy in contact with eachother" & vbCrLf & vbCrLf & _
        "This chatbox is made in VB and will stimulance to continue using VB if you see what you can do with it" & vbCrLf & vbCrLf & _
        "You will have chatboxes, possibility to create your own rooms so you can ask other programmers what you want to know, and a lot more. read about it on the site" & vbCrLf & vbCrLf & _
        "Its totally new and we are still seeking for people. So please, be so kind to check out the program. You will like it mate"
End Sub

Private Sub Label1_Click()
On Error Resume Next
    ShellExecute 0&, "OPEN", "www.bizzit.com", "", "", vbNormalFocus
End Sub

Private Sub rchMainText_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim txt As String

    txt = RichWordOver(rchMainText, X, y)
    If lblCurrentWord.Caption <> txt Then lblCurrentWord.Caption = txt
End Sub

