Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel

''' 
''' Just a wrapper for a 2d array. As long as the number of comments does not exceed x000, this will be sufficiently fast
''' All search methods just perform a linear filter. If comment store grows, consider indexing rows in a hashmap.
Public Class CommentStore
    Public coms As List(Of List(Of String))
    Private excelApp As excel.Application

    Public Sub New()
        Try
            excelApp = New excel.Application
        Catch
            MessageBox.Show("Excel is not installed- it is required for this application to run. Closing...")
            Environment.Exit(0)
        End Try

        coms = New List(Of List(Of String))
        refresh()
    End Sub

    'clear and refresh all editor types
    Public Sub refresh()
        coms.Clear()
        Dim ePaths As List(Of String) = get_editor_file_paths()

        For Each ePath As String In ePaths
            If (Common.verify_editor_file_extension(ePath)) Then
                extract_load_comments(ePath, get_editor_role(ePath))
            End If
        Next
    End Sub

    'overload - refresh a specific editor, looking for its path in the db directory
    Public Function refresh(eRole As String)
        Dim ePaths As List(Of String) = get_editor_file_paths()

        For Each ePath As String In ePaths
            Dim eRoleCand = get_editor_role(ePath)
            If eRoleCand.Equals(eRole) And Common.verify_editor_file_extension(ePath) Then
                clearByEditor(eRole)
                Return extract_load_comments(ePath, eRole)
            End If
        Next

        Return -1
    End Function

    'overload- refresh the store with all new values from the designated file todo should we validate the path as leading to the db dir?
    Public Function refresh(ePath As String, eRole As String)
        If Common.verify_editor_file_extension(ePath) And get_editor_role(ePath).Equals(eRole) Then
            clearByEditor(eRole)
            Return extract_load_comments(ePath, eRole)
        End If
        Return -1
    End Function

    'caller is responsible to delete db file, this just clears the editor in memory!
    Public Sub removeEditor(eRole As String)
        clearByEditor(eRole)
    End Sub

    Private Sub clearByEditor(eRole As String)
        'the old iterate backwards trick to leave indices unperturbed- still O(N^2), could use a hashmap with provider keys?
        For comix As Integer = coms.Count() - 1 To 0 Step -1
            Dim com = coms(comix)
            If com(Common.comStoreIndices.EditorRole).Equals(eRole) Then
                coms.Remove(com)
            End If
        Next

    End Sub

    Private Function get_editor_file_paths() As List(Of String)
        If (Common.FolderExists(Common.DB_PATH)) Then
            Dim eTypeFiles As Array = Directory.GetFiles(Common.DB_PATH)
            Return Common.filterTempFiles(eTypeFiles)
        Else
            'todo log it / messagebox to alert user to contact support
            Return New List(Of String)
        End If

    End Function

    'returns the number of comments extracted. -1 for some error condition
    Private Function extract_load_comments(ePath As String, eRole As String) As Integer
        excelApp.DisplayAlerts = False
        excelApp.Visible = False
        Try
            excelApp.Workbooks.Open(ePath, True) ' open readonly

            Dim eSheet As excel.Worksheet = excelApp.Sheets(1)

            Dim n = eSheet.UsedRange.Rows.Count
            Dim m = eSheet.UsedRange.Columns.Count
            If m = Common.nComProps Then

                '' adding the comments to the store
                For row As Integer = 1 To n
                    Dim newCom As List(Of String) = New List(Of String) From {"", "", "", ""}

                    newCom(Common.comStoreIndices.EditorRole) = eRole
                    newCom(Common.comStoreIndices.Section) = eSheet.Cells(row, Common.dbIndices.Section).Value.ToString
                    newCom(Common.comStoreIndices.CommentName) = eSheet.Cells(row, Common.dbIndices.CommentName).Value.ToString
                    newCom(Common.comStoreIndices.CommentContent) = eSheet.Cells(row, Common.dbIndices.CommentContent).value.ToString

                    coms.Add(newCom)
                Next

                excelApp.Workbooks.Close()

                Return n
            Else
                excelApp.Workbooks.Close()
                Return -1
            End If
        Catch
            'error likely came from opening a workbook that was already open elsewhere.
            MessageBox.Show("Failed to open your comments- do you have the workbook " + ePath + " open already? Close it and try again.")
        End Try
        Return -1
    End Function

    Private Function get_editor_role(ePath As String)
        ''get the editor role
        Dim fields As Array = Common.split_path(ePath)
        Dim fname As String = fields(fields.Length - 1)
        Dim parts As Array = fname.Split(".")
        Dim eRole As String = ""

        If (parts.Length > 1) Then
            For i As Integer = 0 To parts.Length - 2
                eRole += parts(i) + "."
            Next
        End If

        'chop of the trailing .
        If (eRole.Length > 0) Then
            eRole = eRole.Substring(0, eRole.Length - 1)
        End If

        Return eRole
    End Function

    Public Function get_sections(editorRole As String) As List(Of String)
        Dim sects As HashSet(Of String) = New HashSet(Of String)
        For Each com As List(Of String) In coms
            If com(Common.comStoreIndices.EditorRole).Equals(editorRole) Then
                sects.Add(com(Common.comStoreIndices.Section))
            End If
        Next

        Return sects.ToList()
    End Function

    Public Function get_comments(editorRole As String, section As String) As List(Of Comment)
        Dim retval As List(Of Comment) = New List(Of Comment)
        For Each com As List(Of String) In coms
            If (com(Common.comStoreIndices.EditorRole).Equals(editorRole)) Then
                If (com(Common.comStoreIndices.Section).Equals(section)) Then
                    retval.Add(New Comment(com(Common.comStoreIndices.CommentContent),
                                           com(Common.comStoreIndices.CommentName)))
                End If
            End If
        Next
        Return retval
    End Function

    Public Function get_editor_types() As List(Of String)
        Dim eTypes As HashSet(Of String) = New HashSet(Of String)
        For Each com As List(Of String) In coms
            eTypes.Add(com(Common.comStoreIndices.EditorRole))
        Next
        Return eTypes.ToList()
    End Function

End Class