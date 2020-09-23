Attribute VB_Name = "Compare"

Public Function WriteDiffer(ByVal MasterFile As String, ByVal SourceFile As String, ByVal NewFile As String) As String
    ' NOTE: Microsoft Scripting runtime(scrrun.dll)needed

    Dim oFileSystem As FileSystemObject
    Dim oNewFile As TextStream
    Dim oMasterFile As TextStream
    Dim oSourceFile As TextStream
    Set oFileSystem = New FileSystemObject
    Set oMasterFile = oFileSystem.OpenTextFile(MasterFile, ForReading)
    Set oSourceFile = oFileSystem.OpenTextFile(SourceFile, ForReading)
    Set oNewFile = oFileSystem.OpenTextFile(NewFile, ForAppending, True)
    sMasterFile = oMasterFile.ReadAll
    Dim LineNum As Integer


    Do Until oSourceFile.AtEndOfStream
        LineNum = LineNum + 1
        sCurrent = oSourceFile.ReadLine


        If InStr(1, sMasterFile, sCurrent) Then
            ' Line currently already exists in MasterFile (ignore it)
        Else
            oNewFile.WriteLine "(" & LineNum & ")" & Space(3) & sCurrent ' Write what does not exists in MasterFile
        End If
    Loop
    oNewFile.Close
    oMasterFile.Close
    oSourceFile.Close
    Set oSourceFile = Nothing
    Set oMasterFile = Nothing
    Set oNewFile = Nothing
    Set oFileSystem = Nothing
End Function



