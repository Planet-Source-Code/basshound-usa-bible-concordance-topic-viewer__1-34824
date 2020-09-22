Attribute VB_Name = "Module2"
Public Function FileExist(asPath As String) As Boolean
    
    If UCase(Dir(asPath)) = UCase(TrimPath(asPath)) Then
      FileExist = True
    Else
      FileExist = False
    End If
End Function

Public Function TrimPath(ByVal asPath As String) As String
    
    If Len(asPath) = 0 Then Exit Function
    Dim x As Integer
    
    Do
        x = InStr(asPath, "\")
        If x = 0 Then Exit Do
        asPath = Right(asPath, Len(asPath) - x)
    Loop
    TrimPath = asPath
End Function
