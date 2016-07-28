Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO

Public Class ChimeraHome

    Public oDoc As Word.Document
    Public oWord As Word.Application
    Public cStore As CommentStore

    '
    ' start of event handlers
    '

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''todo thread this, make sure the thread has completed before exiting.
        cStore = New CommentStore()

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

    'open the word document and open the editor form
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "~"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim filep As String = openFileDialog1.FileName

            'open a new form corresponding to the selected editor type
            Dim editorForm As ChimeraEditor = New ChimeraEditor(cStore, filep, Me)
            editorForm.Show()
            Me.Hide()

        End If
    End Sub

    Private Sub Form1_Close(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        'noop
    End Sub

    'for editting and adding comments through excel
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim cEditorForm As CommentEditorHome = New CommentEditorHome(Me, cStore)
        cEditorForm.Show()
        Me.Hide()
    End Sub

End Class