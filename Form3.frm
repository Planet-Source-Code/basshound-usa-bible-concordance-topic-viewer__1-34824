VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   1695
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   1095
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add Verse"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Scripture"
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Current Verse"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Versus"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chapter"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Book"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If ScriptOne <> True Then
If wf = 1 Then
If Form2.Text3.Text > "" Then

Text6 = Text6 & "<!--new-->" & vbNewLine
End If
End If


If wf = 2 Then
If Form4.Text2.Text > "" Then

Text6 = Text6 & "<!--new-->" & vbNewLine
End If
End If
Text6 = Text6 & "<!--book-->" & Text1 & vbNewLine & "<!--Chapter-->" & Text2 & vbNewLine & "<!--versus-->" & Text3 & vbNewLine & "<!--verse-->" & Text4 & vbNewLine & "<!--scripture-->" & Text5 & vbNewLine
ScriptOne = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4 = ""
Text5 = ""
Text4.SetFocus
Else
Text6 = Text6 & "<!--verse-->" & Text4 & vbNewLine & "<!--scripture-->" & Text5 & vbNewLine
Text4 = ""
Text5 = ""
Text4.SetFocus
End If

Text6 = Text6 & Text4 & vbNewLine & Text5 & vbNewLine
End Sub

Private Sub Command2_Click()
If wf = 1 Then
Form2.Text3.Text = Form2.Text3.Text & Form3.Text6.Text

End If
If wf = 2 Then
Form4.Text2.Text = Form4.Text2.Text & Form3.Text6.Text
End If

Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
ScriptOne = False
End Sub

