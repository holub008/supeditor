Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO


Public Class ChimeraEditor
    'todo separate filenames from text to be displayed on the document
    'todo always save a backup copy in case the form botches up the document

    Private oWord As Word.Application
    Private oDoc As Word.Document
    Private eStore As CommentStore
    Private origin As Form


    Private keyTag = 55455 'magic identifier for dynamically generated buttons needing later deletion

    Public Sub New(e As CommentStore, fPath As String, orig As Form)

        ' This call is required by the designer?
        InitializeComponent()

        eStore = e
        origin = orig

        '
        ' prep the word doc
        '
        'create the MS word objects to open the file
        oWord = CreateObject("Word.Application")
        oWord.Visible = True


        'todo handle the document was already open and failed to open
        Try
            oDoc = oWord.Documents.Open(fPath)
        Catch
            'fPath doc may already be open for editing- yikes!
            MessageBox.Show("Failed to open your document. Please close all word documents & try again or contact support.")
            Me.Close()
        End Try
        '
        ' building the form
        '
        For Each eType As String In eStore.get_editor_types()
            ComboBox1.Items.Add(eType)
        Next
    End Sub

    Public Sub clear_comments()
        For i As Integer = Panel1.Controls.Count - 1 To 0 Step -1
            Dim ctrl = Panel1.Controls(i)

            If TypeOf (ctrl) Is Button And Not (ctrl.Tag Is Nothing) And ctrl.Tag = 55455 Then
                ctrl.Dispose()
            End If
        Next
    End Sub

    Public Sub populate_comments()
        'first remove any prior comments
        clear_comments()

        'check to see if anything is selected
        If (Not (ComboBox1.SelectedItem Is Nothing) And Not (ComboBox2.SelectedItem Is Nothing)) Then
            Dim eType = ComboBox1.SelectedItem.ToString
            Dim sec = ComboBox2.SelectedItem.ToString

            Dim row = 5
            Dim rowJump = 35
            Dim col1 = 5
            Dim col2 = 170
            Dim col = col1
            Dim butWidth = 155
            Dim butHeight = 25
            Dim SMALL_BUTTON_COMMENT_SIZE = 25

            'comment names that overflow buttons end up on bottom
            'todo subroutine for inserting buttons
            Dim longComments As List(Of Comment) = New List(Of Comment)

            For Each com As Comment In eStore.get_comments(eType, sec)
                If (Len(com.name) <= SMALL_BUTTON_COMMENT_SIZE) Then
                    Dim tempbut = New Button()

                    ''button click & mouseover handlers for comments
                    ' note lambda functions maintain scope variables
                    AddHandler tempbut.Click, Sub()
                                                  insert_comment(com.comment)
                                              End Sub
                    AddHandler tempbut.MouseHover, Sub()
                                                       display_comment(com.comment)
                                                   End Sub
                    AddHandler tempbut.MouseLeave, Sub()
                                                       display_comment("")
                                                   End Sub

                    'button data bindings
                    tempbut.Text = com.name
                    tempbut.Tag = keyTag
                    tempbut.Location = New System.Drawing.Point(col, row)
                    tempbut.Size = New System.Drawing.Size(butWidth, butHeight)

                    Panel1.Controls.Add(tempbut)
                    tempbut.BringToFront()
                    tempbut.BackColor = Color.Lavender

                    'change the button position TODO modular magic :)
                    If (col = col2) Then
                        row = row + rowJump
                        col = col1
                    Else
                        col = col2
                    End If
                Else
                    longComments.Add(com)
                End If
            Next

            'now insert the longer comments as a single button per row
            If (col = col2) Then
                row = row + rowJump
                col = col1
            End If

            butWidth = col2 + butWidth - col1

            For Each com As Comment In longComments

                Dim tempbut = New Button()

                ''button click & mouseover handlers for comments
                ' note lambda functions maintain scope variables
                AddHandler tempbut.Click, Sub()
                                              insert_comment(com.comment)
                                          End Sub
                AddHandler tempbut.MouseHover, Sub()
                                                   display_comment(com.comment)
                                               End Sub

                AddHandler tempbut.MouseLeave, Sub()
                                                   display_comment("")
                                               End Sub

                'button data bindings
                tempbut.Text = com.name
                tempbut.Tag = keyTag
                tempbut.Location = New System.Drawing.Point(col, row)
                tempbut.Size = New System.Drawing.Size(butWidth, butHeight)

                Panel1.Controls.Add(tempbut)
                tempbut.BringToFront()
                tempbut.BackColor = Color.Lavender 'todo weird coloration of button border- close to lavender

                'change the button position
                row = row + rowJump
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
            oWord.Quit()
        Catch
            'nothing to do- the doc is already closed. or we have a null reference, which is really unexpected :/
        End Try

        'display the main form to open other docs
        origin.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'first, we must repopulate the dropdown
        'note selecteditem isn't null since it changed
        ComboBox2.Items.Clear()
        For Each sect As String In eStore.get_sections(ComboBox1.SelectedItem)
            ComboBox2.Items.Add(sect)
        Next
        'next, repopulate the comments

        populate_comments()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        populate_comments()
    End Sub


    '
    ' helpers
    '
    Private Sub insert_comment(com As String)
        Try
            oDoc.Content.Comments.Add(oDoc.ActiveWindow.Selection.Range, com)
        Catch
            MsgBox("The document has been closed- please close your editing session and restart")
        End Try
    End Sub

    Private Sub display_comment(com As String)
        TextBox3.Text = com
    End Sub

End Class