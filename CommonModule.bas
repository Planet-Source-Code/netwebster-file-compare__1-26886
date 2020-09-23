Attribute VB_Name = "Common"
Public Sub FileAppend(FromPath As String, ToPath As String)
   
   Dim Fromno As Integer
   Dim ToNo As Integer
   Dim LineIn As String
   Fromno = FreeFile
   Call fileopen(FromPath, "INPUT", "SHARED", Fromno, 0)
   ToNo = FreeFile
   Call fileopen(ToPath, "APPEND", "LOCKED", ToNo, 0)
   While Not EOF(Fromno)
      Line Input #Fromno, LineIn
      Print #ToNo, LineIn
   Wend
   Close #Fromno, #ToNo

End Sub

Sub filekill(FileSpec As String)
   On Local Error GoTo fkerror
   Kill FileSpec
   On Local Error GoTo 0
   Exit Sub
fkerror:
   Resume Next
End Sub

Sub filemove(Fromspec As String, ToSpec As String)
    On Local Error GoTo fmerror
    Name Fromspec As ToSpec
    On Local Error GoTo 0
    Exit Sub
fmerror:
    Resume Next

End Sub

Function fileopen(FileSpec As String, Method As String, Access As String, fileno As Integer, FileLength As Integer) As Integer
  
  Dim ErrorCode As Integer
  Dim Count As Integer
  Err.Clear
  'On Local Error GoTo foerror
    Do
      ErrorCode = 0
      If UCase$(Method$) = "INPUT" Then
         If Left$(UCase$(Access$), 4) <> "LOCK" Then
            Open FileSpec$ For Input Shared As #fileno
         Else
            Open FileSpec$ For Input Lock Read Write As #fileno
         End If
      ElseIf UCase$(Method$) = "APPEND" Then
         Open FileSpec$ For Append Lock Read Write As #fileno
      ElseIf UCase$(Method$) = "OUTPUT" Then
         Open FileSpec$ For Output Lock Read Write As #fileno
      ElseIf UCase$(Method$) = "BINARY" Then
         Open FileSpec$ For Binary Lock Read Write As #fileno
      Else
         If UCase$(Access$) = "SHARED" Then
            Open FileSpec$ For Random Shared As #fileno Len = FileLength
         Else
            Open FileSpec$ For Random Lock Read Write As #fileno Len = FileLength
         End If
      End If
      If ErrorCode = 70 Then
        Count = Count + 1
        If Count = 30 Then
            Count = 0: Beep
        End If
      End If
    Loop While ErrorCode = 70
    On Local Error GoTo 0
    fileopen = ErrorCode
End Function

Public Sub Main()
    Load frmGetText
End Sub

Public Sub UnloadAllForms(Optional sFormName As String = "")

   Dim Form As Form

   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form

End Sub

Sub CenterForm(FormName As Form, Optional deltaX As Integer, Optional deltaY As Integer)
    FormName.Left = (Screen.Width - FormName.Width) / 2 - deltaX
    FormName.Top = (Screen.Height - FormName.Height) / 2 - deltaY
End Sub

Function exists(FileSpec As String) As Boolean
   
   Dim ErrorCode As Integer
   Dim X As String
   ErrorCode = 0
   On Local Error GoTo exerror
   X = Dir(FileSpec)
   On Local Error GoTo 0
   If X$ = "" Or ErrorCode <> 0 Then
      exists = False
   Else
      exists = True
   End If
   Exit Function

exerror:
   ErrorCode = Err
   Resume Next
End Function

