VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notepad"
   ClientHeight    =   4710
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Save 
         Caption         =   "&SaveAs"
      End
      Begin VB.Menu Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu Minimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Cut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Copy_Click()
Text1.SelStart = Text1.SelLength
Clipboard.Clear
   Clipboard.SetText Text1.Text
End Sub
Private Sub Cut_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
End Sub
Private Sub Exit_Click()
On Error GoTo ErrorHandler
Dim Msg, Style, Title, Response, MyString
Msg = "Are you sure you want to exit ?"
Style = vbYesNo + vbQuestion + vbDefaultButton1
Title = "Warning"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
   MyString = "Yes"
End
End If
ErrorHandler:
End Sub
Private Sub Minimize_Click()
Form1.WindowState = 1
End Sub
Private Sub New_Click()
Text1.Text = ""
End Sub
Private Sub Open_Click()
   CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
   CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
  Dim LoadFileToTB As Boolean
 Dim TxtBox As Object
 Dim FilePath As String
  Dim Append As Boolean
Dim iFile As Integer
Dim s As String
If Dir(FilePath) = "" Then Exit Sub
On Error GoTo ErrorHandler:
s = Text1.Text
iFile = FreeFile
Open CommonDialog1.FileName For Input As #iFile
s = Input(LOF(iFile), #iFile)
If Append Then
    Text1.Text = Text1.Text & s
Else
    Text1.Text = s
End If
LoadFileToTB = True
ErrorHandler:
If iFile > 0 Then Close #iFile
End Sub
Private Sub Paste_Click()
Text1.SelText = Clipboard.GetText()
End Sub
Private Sub Print_Click()
 On Error GoTo ErrHandler
  Dim BeginPage, EndPage, NumCopies, i
   CommonDialog1.CancelError = True
  CommonDialog1.ShowPrinter
  BeginPage = CommonDialog1.FromPage
  EndPage = CommonDialog1.ToPage
  NumCopies = CommonDialog1.Copies
  For i = 1 To NumCopies
 Printer.Print Text1.Text
  Next i
  Exit Sub
ErrHandler:
   Exit Sub
End Sub
Private Sub Save_Click()
On Error GoTo ErrorHandler
  CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 2
   CommonDialog1.ShowSave
 CommonDialog1.FileName = CommonDialog1.FileName
Dim iFile As Integer
 Dim SaveFileFromTB As Boolean
 Dim TxtBox As Object
 Dim FilePath As String
Dim Append As Boolean
  iFile = FreeFile
If Append Then
    Open CommonDialog1.FileName For Append As #iFile
Else
    Open CommonDialog1.FileName For Output As #iFile
End If
Print #iFile, Text1.Text
SaveFileFromTB = True
ErrorHandler:
Close #iFile
End Sub

