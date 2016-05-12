Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO

Public Class supeditor_selector

    Public oDoc As Word.Document
    Public oWord As Word.Application
    Public editorTypes As List(Of String)
    Public APPDATA_DIR As String = "\supedit_editor\"
    Public COMMENTSTORE_FORMAT = ".tsv"

    '
    ' start of event handlers
    '

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'place the different editing options into the listbox
        editorTypes = get_editor_types()
        For Each eType As String In editorTypes
            ListBox1.Items.Add(eType)
        Next

        If (editorTypes.Count() > 0) Then
            ListBox1.SelectedIndex = 0
        Else
            MessageBox.Show("You have a bad install. Contact support with error: Could not find editor type files.")
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "~"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim tester = openFileDialog1.FileName

            'create the MS word objects to open the file
            oWord = CreateObject("Word.Application")
            oWord.Visible = True
            'todo handle the document was already open and failed to open
            Try
                oDoc = oWord.Documents.Open(tester)

                'open a new form corresponding to the selected editor type
                Dim editorForm As GenericEditorForm
                'todo pass pass oword for closing application and not just doc
                editorForm = New GenericEditorForm(ListBox1.SelectedItem, oDoc)
                editorForm.Show()
                Me.Hide()
            Catch
                'do nothing
            End Try


        End If
    End Sub

    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        'todo some better way to handle this than just saving!
        Try
            oDoc.Save()
            oDoc.Close()

        Catch
            'nothing to do- the doc has been manually closed.
        End Try

    End Sub

    '
    ' start of helper routines - not event handlers
    '

    Private Function get_editor_types() As List(Of String)
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

End Class