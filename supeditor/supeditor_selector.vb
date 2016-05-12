Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO

Public Class supeditor_selector

    Public oDoc As Word.Document
    Public oWord As Word.Application

    '
    ' start of event handlers
    '

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'warn users if they have a currently open word document
        For Each b As Process In Process.GetProcesses(".")
            Try
                If b.MainWindowTitle.Length > 0 Then
                    If (b.ProcessName.ToString() = "WINWORD") Then
                        TextBox1.Text = "It looks like you currently have a word document open. If the document you wish to edit is currently open, please close it before proceeding."
                        TextBox1.BorderStyle = BorderStyle.Fixed3D
                        TextBox1.Enabled = True
                    End If
                End If
            Catch
            End Try
        Next

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "~"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim filep = openFileDialog1.FileName

            'create the MS word objects to open the file
            oWord = CreateObject("Word.Application")
            oWord.Visible = True

            'todo handle the document was already open and failed to open
            ' Try
            oDoc = oWord.Documents.Open(filep)

            'open a new form corresponding to the selected editor type
            Dim editorForm As GenericEditorForm
            'todo pass pass oword for closing application and not just doc
            editorForm = New GenericEditorForm(oDoc)
            editorForm.Show()
            Me.Hide()
            'Catch
            'MessageBox.Show("Oops: failed to open the editor- please try again! ")
            'End Try


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
    ' helper routines - not event handlers
    '

End Class