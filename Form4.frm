VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Scripure"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   645
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Open Curpath & "\" + "des.dat" For Output As #1

Print #1, Text1.Text
Close #1


Open Curpath & "\" + "des.html" For Output As #1

Print #1, Text1.Text
Close #1

Open Curpath & "\" + "script.dat" For Output As #1

Print #1, Text2.Text
Close #1

Text3 = Text3 & "<table bgcolor=""lightgrey""><tr><td>"
Open Curpath & "\" + "script.dat" For Input As #1
CountStuff = 0
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

Open Curpath & "script.html" For Output As #1

Print #1, Text3.Text
Close #1
Unload Me
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
wf = 2

Close #1
Open Curpath & "des.dat" For Input As #1

Do Until EOF(1)
Line Input #1, textline

All = All & textline & vbNewLine

Loop

Close #1
Text1 = All

Close #1
All = ""
Open Curpath & "script.dat" For Input As #1

Do Until EOF(1)
Line Input #1, textline

All = All & textline & vbNewLine

Loop

Close #1
Text2 = All
End Sub

Private Sub Form_Unload(Cancel As Integer)
wf = 1
End Sub

