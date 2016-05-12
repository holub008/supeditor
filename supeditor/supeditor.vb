Imports Word = Microsoft.Office.Interop.Word
Imports System


Public Class GenericEditorForm
    'todo separate filenames from text to be displayed on the document
    'todo always save a backup copy in case the form botches up the document

    Public eType As String
    Public oDoc As Word.Document
    Public comStore As CommentStore

    Public Sub New(e As String, o As Word.Document)

        ' This call is required by the designer?
        InitializeComponent()

        eType = e
        oDoc = o

        '
        ' building the form
        '

        'ListBox1.DrawMode = DrawMode.OwnerDrawFixed

        comStore = New CommentStore(eType)

        ' first add radio buttons for each subsection type
        Dim subSecs As List(Of String) = comStore.get_subsections()

        For Each subsec As String In subSecs
            Dim newButton As New ToolStripMenuItem

            With newButton
                .Text = subsec
            End With

            'when the section is clicked, load in the correct comments
            AddHandler newButton.Click, AddressOf SectButtonHandler

            MenuStrip1.Items.Add(newButton)

        Next


    End Sub

    '
    ' handler routines
    '

    Private Sub SectButtonHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'first remove the old list items
        ListBox1.Items.Clear()

        'now insert the comments for the given section
        Dim sec As String = CType(CType(sender, System.Windows.Forms.ToolStripMenuItem).Text, String)
        For Each comcon In comStore.get_comments(sec)
            ListBox1.Items.Add(comcon.comment)
        Next

    End Sub

    'straight jacked from https://social.msdn.microsoft.com/Forums/en-US/3dee72ea-83d9-4a59-95d0-ac1b93432b11/listbox-with-alternate-row-colors?forum=vbgeneral
    'todo switching stripmenu with an item selected busts the bounds -try moving the highlighting logic to the selectedindexchange henderl
    Private Sub ListBox1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ListBox1.DrawItem
        If e.Index Mod 2 = 0 Then
            e.Graphics.FillRectangle(Brushes.LightGray, e.Bounds)
        End If

        e.Graphics.DrawString(ListBox1.Items(e.Index).ToString, Me.Font, Brushes.Black, 0, e.Bounds.Y + 2)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        'ListBox1.Refresh()
    End Sub


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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim currentComment As String = ListBox1.SelectedItem
            If (Not currentComment = "") Then
                oDoc.Content.Comments.Add(oDoc.ActiveWindow.Selection.Range, currentComment)
            End If
        Catch
            MsgBox("The document has been closed- please close your editing session and restart")
        End Try
    End Sub

    '
    ' helper functions
    '
End Class

Public Class CommentStore
    Public lookup As New Dictionary(Of String, List(Of CommentContainer))
    Public eType As String

    'todo vb equivalent of a header to share these guys between form 1 and 2?
    Public APPDATA_DIR As String = "\supedit_editor\"
    Public COMMENTSTORE_FORMAT = ".tsv"

    Public Sub New(e As String)
        eType = e

        ''preparing the lookup structure to store comments
        'commentData is a tab delimited strings (newlines for entries)
        Dim commentData As String = My.Computer.FileSystem.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & APPDATA_DIR & eType & COMMENTSTORE_FORMAT)
        Dim commentEntries As Array = commentData.Split(vbNewLine)

        For Each commentEntry As String In commentEntries
            Dim fields As Array = commentEntry.Split(vbTab)

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