
Module Common
    Public Enum comStoreIndices
        EditorRole = 0
        Section = 1
        CommentName = 2
        CommentContent = 3
    End Enum

    Public Enum dbIndices
        Section = 1
        CommentName = 2
        CommentContent = 3
    End Enum


    Public nComProps As Integer = 3

    Public DB_PATH As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Chimera\"

    Public excelDefaultSheetName As String = "sheet1"

    Public AddHelp As String = "The excel sheet contains rows that represent comments. Each row has 3 columns, in order: Section, Comment Name, and Comment. If you want a line break to appear in your comment, insert these characters: '\n'"

    Public comStoreExtensions As List(Of String) = New List(Of String) From {".xlsx", ".xls"}

    'check to see if this is an 
    Public Function verify_editor_file_extension(fpath As String) As Boolean
        Dim steps As Array = split_path(fpath)
        Dim fileName As String = steps(steps.Length - 1)

        'verify that we have an extension excel can use
        Dim nameParts As Array = fileName.Split(".")
        Dim ext As String = "." & nameParts(nameParts.Length - 1)
        If (comStoreExtensions.Contains(ext)) Then
            Return True
        End If

        Return False

    End Function

    Public Function get_extension(fpath As String) As String
        Dim fields As Array = fpath.Split(".")
        If fields.Length > 1 Then
            Return fields(fields.Length - 1)
        End If
        Return ""
    End Function

    Public Function split_path(fpath As String) As Array
        Return fpath.Split("\")
    End Function

    Public Function FolderExists(strFolderPath As String) As Boolean
        On Error Resume Next
        FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
        On Error GoTo 0
    End Function

    Public Function copy_file(fromPath As String, toPath As String)
        'note this will rename the file as well todo enum for return conditions
        If (Not My.Computer.FileSystem.FileExists(fromPath)) Then
            Return -2
        End If
        If (Not My.Computer.FileSystem.FileExists(toPath)) Then
            My.Computer.FileSystem.CopyFile(fromPath, toPath)
            Return 1
        End If
        Return -1
    End Function

    Public Function delete_file(fPath As String)
        If (Not My.Computer.FileSystem.FileExists(fPath)) Then
            Return -1
        End If
        My.Computer.FileSystem.DeleteFile(fPath)
        Return 1
    End Function

    Public Function filterTempFiles(PathsIn As Array) As List(Of String)
        Dim ret As List(Of String) = New List(Of String)
        For Each path As String In PathsIn
            Dim fields As Array = path.Split("\")
            If Not fields(fields.Length - 1).Substring(0, 2) = "~$" Then
                ret.Add(path)
            End If
        Next
        Return ret
    End Function

End Module