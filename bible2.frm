VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add Scripture"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Topic"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Type additional notes here:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "New topic Name"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Then
MsgBox "You must have a topic title!", vbInformation, "Unable to comply..."
Exit Sub
End If
Close #1
MkDir Curpath + Text1.Text + "\"
Open Curpath + "list.dat" For Append As #1

Print #1, Text1.Text
Close #1


Open Curpath + Text1.Text + "\" + "des.dat" For Append As #1

Print #1, Text2.Text
Close #1


Open Curpath + Text1.Text + "\" + "des.html" For Append As #1

Print #1, Text2.Text
Close #1

Open Curpath + Text1.Text + "\" + "script.dat" For Append As #1

Print #1, Text3.Text
Close #1

CountStuff = 0
Text3 = ""
Text3 = Text3 & "<table bgcolor=""lightgrey""><tr><td>"
Open Curpath + Text1.Text + "\" + "script.dat" For Input As #1

Do Until EOF(1)
CountStuff = CountStuff + 1
Line Input #1, lineoftext
If lineoftext = "<!--new-->" Then
Text3 = Text3 & "</td></tr></table><br><br>"
Text3 = Text3 & "<table bgcolor=""lightgrey""><tr><td>"

CountStuff = 10
End If

If CountStuff = 1 Then
Text3 = Text3 & "<b>Book: </b>" & lineoftext & "</td></tr>"
ElseIf CountStuff = 2 Then
Text3 = Text3 & "<tr><td><b>Chapter: </b>" & lineoftext & "</td></tr>"
ElseIf CountStuff = 3 Then
Text3 = Text3 & "<tr><td><b>Versus: </b>" & lineoftext & "</td></tr></table><br><table bgcolor=""lightgrey""><tr><td>"
ElseIf CountStuff = 4 Then
Text3 = Text3 & lineoftext & "<div style=""margin-left : 40px;"" >"
ElseIf CountStuff = 5 Then
Text3 = Text3 & lineoftext & "</div>"
CountStuff = 3
ElseIf CountStuff = 10 Then
CountStuff = 0
End If



Loop
Close #1
Text3 = Text3 & "</td></tr></table>"

Open Curpath + Text1.Text + "\" + "script.html" For Append As #1

Print #1, Text3.Text
Close #1

Form1.List1.AddItem Text1.Text
Unload Me
End Sub


Private Sub Command2_Click()
Form3.Show
End Sub

