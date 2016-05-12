Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO


Public Class GenericEditorForm
    'todo separate filenames from text to be displayed on the document
    'todo always save a backup copy in case the form botches up the document

    Public eType As String
    Public oDoc As Word.Document
    Public eStore As EditorStore

    Public Sub New(o As Word.Document)

        ' This call is required by the designer?
        InitializeComponent()

        oDoc = o
        eStore = New EditorStore()

        '
        ' building the form
        '
        For Each eType As String In eStore.get_editor_types()
            ComboBox1.Items.Add(eType)
        Next


    End Sub

    Public Sub populate_comments()
        TextBox3.Text = ""
        If (Not (ComboBox1.SelectedItem Is Nothing) And Not (ComboBox2.SelectedItem Is Nothing)) Then
            Dim eType = ComboBox1.SelectedItem.ToString
            Dim sec = ComboBox2.SelectedItem.ToString

            For Each com As CommentContainer In eStore.get_comments(eType, sec)
                TextBox3.Text = TextBox3.Text & com.name
            Next
        End If

        'If nothing is selected- no problem.

    End Sub

    '
    ' handler routines
    '


    Private Sub GenericEditorForm_Close(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        'todo some better way to handle this than just saving!
        Try
            oDoc.Save()
            oDoc.Close()
        Catch
            'nothing to do- the doc has been manually closed.
        End Try

        'display the main form to open other docs
        supeditor_selector.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Try
            Dim currentComment = "blah"
            oDoc.Content.Comments.Add(oDoc.ActiveWindow.Selection.Range, currentComment)
        Catch
            MsgBox("The document has been closed- please close your editing session and restart")
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'first, we must repopulate the dropdown
        'note selecteditem isn't null since it changed
        ComboBox2.Items.Clear()
        For Each sect As String In eStore.get_subsections(ComboBox1.SelectedItem)
            ComboBox2.Items.Add(sect)
        Next
        'next, repopulate the comments

        populate_comments()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        populate_comments()
    End Sub
End Class

Public Class EditorStore
    Public lookup As New Dictionary(Of String, CommentStore)
    'todo vb equivalent of a header to share these guys?
    Public APPDATA_DIR As String = "\supedit_editor\"
    Public COMMENTSTORE_FORMAT = ".tsv"

    Public Sub New()
        Dim eTypes = get_editor_types_from_appdata()
        For Each eType As String In eTypes
            lookup.Add(eType, New CommentStore(eType))
        Next
    End Sub

    Private Function get_editor_types_from_appdata() As List(Of String)
        'todo cast to list of string instead of ugly copy?
        'todo exception handling (specifically, what if non-tsv files here?
        Dim eTypeDirs As Array = Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & APPDATA_DIR)

        Dim retval As List(Of String) = New List(Of String)
        For Each temp As String In eTypeDirs
            Dim fields As Array = temp.Split("\")
            Dim fname As String = fields(fields.Length - 1)
            retval.Add(fname.Substring(0, fname.IndexOf(COMMENTSTORE_FORMAT)))
        Next

        Return (retval)
    End Function

    Public Function get_subsections(editorType As String)
        Return (lookup(editorType).get_subsections())
    End Function

    Public Function get_comments(editorType As String, section As String)
        Dim eComStore = lookup(editorType)
        Return (eComStore.get_comments(section))
    End Function

    Public Function get_editor_types()
        Dim keyList As New List(Of String)
        For Each key As String In lookup.Keys
            keyList.Add(key)
        Next

        Return (keyList)
    End Function
End Class

Public Class CommentStore
    Public lookup As New Dictionary(Of String, List(Of CommentContainer))
    Public eType As String

    'todo vb equivalent of a header to share these guys?
    Public APPDATA_DIR As String = "\supedit_editor\"
    Public COMMENTSTORE_FORMAT = ".tsv"
    Public COMMENTSTORE_DELIM = vbTab

    Public Sub New(e As String)
        eType = e

        ''preparing the lookup structure to store comments
        'commentData is a tab delimited strings (newlines for entries)
        Dim commentData As String = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & APPDATA_DIR & eType & COMMENTSTORE_FORMAT)
        Dim commentEntries As Array = commentData.Split(vbNewLine)

        For Each commentEntry As String In commentEntries
            Dim fields As Array = commentEntry.Split(COMMENTSTORE_DELIM)

            'todo bounds checking
            Dim section As String = fields(0)
            'todo wth? must be dos encoding/ me not understanding vb split function
            'remove the leading newline char
            If (Strings.Asc(section) = 10) Then
                section = section.Substring(1)
            End If

            Dim name As String = fields(1)
            Dim comment As String = fields(2)

            If Not lookup.ContainsKey(section) Then
                lookup.Add(section, New List(Of CommentContainer))
            End If

            'todo what to do with the comment name -> a container class for comment name, comment text
            lookup(section).Add(New CommentContainer(comment, name))
        Next

    End Sub

    Public Function get_comments(sec As String)
        If (lookup.ContainsKey(sec)) Then
            Return lookup(sec)
        Else
            Return New List(Of String) 'there's nothing stored for this section
        End If

    End Function

    Public Function get_subsections()
        Dim keyList As New List(Of String)
        For Each key As String In lookup.Keys
            keyList.Add(key)
        Next

        Return (keyList)
    End Function


End Class

'' a struct class for keeping track of 
Public Class CommentContainer

    Public comment As String
    Public name As String

    Public Sub New(c As String, n As String)
        comment = c
        name = n
    End Sub

End Class