Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel


Public Class CommentEditorHome
    Private Shared cStore As CommentStore
    Private origin As Form
    Private excelApp As excel.Application

    Public Sub New(o As Form, c As CommentStore)
        InitializeComponent()

        cStore = c
        origin = o

        excelApp = New excel.Application

        populateForm()

    End Sub

    Private Sub populateForm()
        ComboBox1.Items.Clear()
        For Each eRole As String In cStore.get_editor_types()
            ComboBox1.Items.Add(eRole)
        Next
    End Sub

    'edit the existing comment store
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex > -1 Then
            Dim ePath = Common.DB_PATH + ComboBox1.SelectedItem.ToString

            Dim excelAppForUser As excel.Application
            Try
                excelAppForUser = New excel.Application

                'todo hide the form until the user has finished editing. visual basic is a piece of shit that I don't understand. why shouldn't a mutex be used in an unsycnchronized context. isn't that the fucking point?
                'todo present them with a form explaining what to do before .visible=True
                AddHandler excelAppForUser.WorkbookBeforeClose, Sub()
                                                                    cStore.refresh()
                                                                End Sub

                'open the sheet for the user to edit
                excelAppForUser.DisplayAlerts = True
                excelAppForUser.Workbooks.Open(ePath)
                excelAppForUser.Visible = True
            Catch
                MessageBox.Show("Unable to open Excel for editing comments. Try again or contact support. ")
            End Try
        Else
            MessageBox.Show("You have not selected the editor role which you wish to edit.")
        End If

    End Sub

    'integrate comments from an existing spreadsheet- todo if kallmes want it
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "~"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        If ComboBox1.SelectedIndex > -1 Then
            Dim eRole As String = ComboBox1.SelectedItem.ToString()

            'open a file dialog to select the file, then read in the comments
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim filep = openFileDialog1.FileName

                Dim n As Integer = cStore.refresh(eRole)

                If n > 0 Then
                    MessageBox.Show("Successfully added comments from your sheet to the " & eRole & " editor role.")
                Else
                    MessageBox.Show("Unable to add your comments. Close open excel sheets and try again, or contact support.")
                End If
            End If
        Else
            MessageBox.Show("You have not selected the editor role which you wish to add comments to.")
        End If
    End Sub

    Private Sub CommentEditorHome_Close(sender As Object, e As EventArgs) Handles MyBase.FormClosed
        Try
            excelApp.close()
        Catch
            'nothing to do- app already closed
        End Try

        origin.Show()
    End Sub

    'display help info for updating existing editor roles
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        MessageBox.Show(Common.AddHelp)
    End Sub

    'create a new editor from an existing sheet
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim eRole As String = TextBox1.Text
        If Not eRole = "" Then
            'open a file dialog to find the 
            Dim openFileDialog1 As New OpenFileDialog()

            openFileDialog1.InitialDirectory = "~"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            'open a file dialog to select the file, then read in the comments
            If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim filep As String = openFileDialog1.FileName
                Dim dest As String = Common.DB_PATH & eRole & "." & Common.get_extension(filep)

                Dim ret As Integer = Common.copy_file(filep, dest)
                If (ret = -1) Then
                    MessageBox.Show("Failed to create your editor: editor role already exists")
                ElseIf (ret = -2) Then
                    MessageBox.Show("Failed to create your editor: the comment sheet you selected does not exist")
                Else
                    cStore.refresh(eRole)
                    populateForm()
                    MessageBox.Show("Successfully created your new editor role!")
                End If
            End If

        Else
            MessageBox.Show("Please enter a new editor role before selecting a comment sheet.")
        End If
    End Sub

    'delete an editor role
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.SelectedIndex > -1 Then
            Dim eRole As String = ComboBox1.SelectedItem.ToString
            Dim fPath As String = Common.DB_PATH & eRole & Common.comStoreExtensions(0) 'todo extension assumption!
            Dim ret As Integer = Common.delete_file(fPath)
            If (ret = 1) Then
                cStore.removeEditor(eRole)
                populateForm()
                MessageBox.Show("Successfully deleted the " & eRole & " editor role")
            Else
                MessageBox.Show("Failed to delete your editor role- please retry or contact support.")
            End If
        Else
            MessageBox.Show("No editor role is selected to be deleted.")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

End Class