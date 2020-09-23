Attribute VB_Name = "modFileFunc"
'''''''''''''''''''''''''''''''''''''''''''
' File Function Module
'
'
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Public Function FileExists(strPath As String, strName As String) As Boolean
    ' On ErrorResume Next
    If Dir$(strPath) = strName Then FileExists = True
End Function

Public Function CheckPath(strFolderPath As String)
    ' On ErrorResume Next
    Dim intLen As Long
    intLen = Len(strFolderPath)
    If Mid$(strFolderPath, intLen, 1) = "\" Then
        CheckPath = Left$(strFolderPath, (intLen - 1)) 'removes the "\"
    Else
        CheckPath = strFolderPath
    End If
End Function

Public Function GetFileName(path As String) As String
    ' On ErrorResume Next
    Dim i As Integer
    For i = (Len(path)) To 1 Step -1
        If Mid(path, i, 1) = "\" Then
            GetFileName = Mid(path, i + 1, Len(path) - i + 1)
            Exit For
        End If
    Next
End Function

Public Function GetFileExtension(FileName As String)
    ' On ErrorResume Next
    Dim TempStr As String
    TempStr = Right(FileName, 2)
    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Function
    Else
        TempStr = Right(FileName, 3)
        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Function
        Else
            TempStr = Right(FileName, 4)
            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Function
            Else
                TempStr = Right(FileName, 5)
                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
End Function

Public Function SetBytes(Bytes) As String
    ' On ErrorGoTo UUUerror

    If Bytes >= 1073741824 Then
        SetBytes = Format(Bytes / 1024 / 1024 / 1024, "#0.00") _
        & " GB"
    ElseIf Bytes >= 1048576 Then
        SetBytes = Format(Bytes / 1024 / 1024, "#0.00") & " MB"
    ElseIf Bytes >= 1024 Then
        SetBytes = Format(Bytes / 1024, "#0.00") & " KB"
    ElseIf Bytes < 1024 Then
        SetBytes = Fix(Bytes) & " Bytes"
    End If
    Exit Function
UUUerror:
    SetBytes = "0 Bytes"
End Function


Public Sub CopyFile(Source As String, Destination As String)
    On Error GoTo CopyError
    FileCopy Source, Destination
    Exit Sub
CopyError:
    MsgBox "Error copying File"
    Resume
End Sub

Public Function GetFileSize(FileName) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(FileName)


    If TempStr >= "1024" Then
        'KB
        TempStr = CCur(TempStr / 1024) & "KB"
    Else


        If TempStr >= "1048576" Then
            'MB
            TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function


