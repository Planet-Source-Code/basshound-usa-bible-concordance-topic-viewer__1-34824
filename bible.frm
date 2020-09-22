VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   2775
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   2040
      TabIndex        =   2
      Top             =   3600
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   1680
      Picture         =   "bible.frx":0000
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Image3 
      Height          =   540
      Left            =   1080
      Picture         =   "bible.frx":04E8
      Top             =   120
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   600
      Picture         =   "bible.frx":0A9F
      Top             =   120
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "bible.frx":0F6E
      Top             =   120
      Width           =   360
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu dt 
         Caption         =   "&Delete Topic"
      End
      Begin VB.Menu nt 
         Caption         =   "&New Topic..."
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a program i made for my own personl usage
'and to help others program
'I dotn really care what you do with this program, but
'dont be lame and change it to ur name and call it urs!
'Good luck and God Bless



Private Sub Command2_Click()
Curpath = "C:\bible\"

List1.Clear
If FileExist(App.Path + "\list.dat") Then

Open "C:\bible\list.dat" For Input As #1
Do Until EOF(1)
Line Input #1, lineoftext
List1.AddItem lineoftext
Loop

Close #1

End If
Text3.Text = ""
End Sub






Private Sub about_Click()
Form5.Show
End Sub

Private Sub dt_Click()
Call DeleteTopic
End Sub

Private Sub Form_Load()
Call MakeFolder
wf = 1

Text2 = "All Topics"
Curpath = App.Path + "\Dats\"

If FileExist(Curpath & "list.dat") Then

Open Curpath & "list.dat" For Input As #1
Do Until EOF(1)
Line Input #1, lineoftext
List1.AddItem lineoftext
Loop

Close #1
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
List1.Height = Form1.Height - 1800
WebBrowser2.Width = Form1.Width - 2200
WebBrowser1.Width = Form1.Width - 2200
WebBrowser2.Height = Form1.Height - 3000
If WebBrowser2.Height > 2775 Then
WebBrowser2.Height = 2775
End If
WebBrowser1.Top = WebBrowser2.Height + 900
WebBrowser1.Height = List1.Height - WebBrowser2.Height + 200

If Form1.Width < 4000 Then
Form1.Width = 4000
End If
If Form1.Height < 4500 Then
Form1.Height = 5450
End If
End Sub

Private Sub Image1_Click()
Dim Undo As String

If Text2 = "All Topics" Then
MsgBox "Your already at the top level. Cannot go any higher!", vbInformation, "Unable to comply..."

Exit Sub
End If
Undo = List2.List(List2.ListCount - 1)
Curpath = List2.List(List2.ListCount - 1)

Text2 = Mid(Curpath, Len(App.Path) + 6, Len(Curpath) - Len(App.Path) - 6)

If Text2 = "\" Then
Text2 = "All Topics"
End If
Close #1
Open (Undo & "list.dat") For Input As #1
alltext = ""
List1.Clear
Do Until EOF(1)
Line Input #1, lineoftext
List1.AddItem lineoftext
Loop
Close #1

List2.RemoveItem (List2.ListCount - 1)
If Text2 = "" Then
Text2 = "All Topics"
End If
End Sub

Private Sub Image2_Click()
Form2.Show
End Sub

Private Sub Image3_Click()
Call DeleteTopic
End Sub

Private Sub Image4_Click()
If Text2 = "All Topics" Then
MsgBox "Please click on a topic first, then click the edit icon!", vbInformation, "Unable to comply..."
Exit Sub
End If


Form4.Show
End Sub

Private Sub List1_DblClick()

Close #1
If Dir(Curpath & List1.Text & "\" & "des.dat") <> "" Then
Open (Curpath & List1.Text & "\" & "des.dat") For Input As #1

Do Until EOF(1)
Line Input #1, lineoftext
alltext = alltext & lineoftext & vbNewLine
Loop
Close #1
End If
Text3 = alltext
List2.AddItem Curpath
CurFile = Curpath & List1.Text & "\" & "list.dat"
Curpath = Curpath & List1.Text & "\"
Text2 = Mid(Curpath, Len(App.Path) + 6, Len(Curpath) - Len(App.Path) - 6)
List1.Clear
If FileExist(CurFile) Then

Open (CurFile) For Input As #1
alltext = ""
Do Until EOF(1)
Line Input #1, lineoftext
List1.AddItem lineoftext
Loop
Close #1




End If

If FileExist(Curpath & "script.html") Then



Open (Curpath & "script.html") For Input As #1

WebBrowser2.Navigate Curpath & "script.html"
End If





If FileExist(Curpath & "des.html") Then


WebBrowser1.Navigate Curpath & "des.html"

End If
End Sub

Private Sub nst_Click()
Curpath = Curpath & List1.Text & "\"
List1.Clear
Form2.Show
End Sub

Private Sub nt_Click()

Form2.Show
End Sub




Public Sub DeleteTopic()


If List1.Text = "" Then
MsgBox "Unable to delete, no topic was selected. Please select a topic before you try to delete!", vbInformation, "Unable to comply..."
Exit Sub
Else
Dim UserResponse
UserResponse = MsgBox("Are you sure you wish to delete the topic """ & List1.Text & """ and all of its contents (if any)?", vbYesNo, "Confirming delete...")
If UserResponse = vbYes Then
DeleteDirectory Curpath & List1.Text & "\"
List1.RemoveItem (List1.ListIndex)
Kill Curpath + "list.dat"
For i = 0 To List1.ListCount
Open Curpath + "list.dat" For Append As #1
If List1.List(i) <> "" Then
Print #1, List1.List(i)
End If
Close #1
Next i
If List1.List(0) = "" Then
Kill Curpath + "list.dat"
End If
End If
End If

End Sub


Private Sub DeleteDirectory(ByVal dir_name As String)
Dim file_name As String
Dim files As Collection
Dim i As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    file_name = Dir$(dir_name & "\*.*", vbReadOnly + vbHidden + vbSystem + vbDirectory)
    Do While Len(file_name) > 0
        If (file_name <> "..") And (file_name <> ".") Then
            files.Add dir_name & "\" & file_name
        End If
        file_name = Dir$()
    Loop

    ' Delete the files.
    For i = 1 To files.Count
        file_name = files(i)
        ' See if it is a directory.
        If GetAttr(file_name) And vbDirectory Then
            ' It is a directory. Delete it.
            DeleteDirectory file_name
        Else
            ' It's a file. Delete it.
           
        
            SetAttr file_name, vbNormal
            Kill file_name
        End If
    Next i

    ' The directory is now empty. Delete it.
    
    RmDir dir_name
End Sub



Public Sub MakeFolder()
On Error GoTo error
MkDir (App.Path & "\Dats")
error:
Exit Sub
End Sub
