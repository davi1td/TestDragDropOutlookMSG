Imports System.Windows.Forms
Imports NetOffice.OutlookApi.Enums
Imports Outlook = NetOffice.OutlookApi

Public Class Form1
    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.All
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim count As Integer = (ListBox1.Items.Count - 1)
        Dim words As String
        ListBox1.SelectedItem = Nothing
        For a = 0 To count
            words = ListBox1.Items.Item(a)
            If InStr(words.ToLower, TextBox2.Text.ToLower) Then
                ListBox1.SelectedItem = words

                If ListBox1.Items.Count >= ListBox1.SelectedIndex + 2 Then
                    MessageBox.Show(ListBox1.Items(ListBox1.SelectedIndex + 2).ToString)
                End If

                Return
                'or display in msgbox Like
                'msgbox(words)
            End If
        Next
    End Sub
    Private Sub ListBox1_DragEnter(sender As Object, e As DragEventArgs) Handles ListBox1.DragEnter
        e.Effect = DragDropEffects.All
    End Sub

    Private Sub ListBox1_DragDrop(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
        Dim app As Outlook.Application = New Outlook.Application()
        Dim appexplorer As Outlook.Explorer = app.ActiveExplorer
        If e.Data.GetDataPresent("FileGroupDescriptor") Then

            For Each testForEmail In appexplorer.Application.ActiveExplorer.Selection()
                If TypeOf testForEmail Is Outlook.MailItem Then
                    ProcessEmail(testForEmail)
                End If
            Next
        End If

    End Sub
    Private Sub ProcessEmail(theEmail As Outlook.MailItem)


        ListBox1.Items.Clear()
        ListBox1.Items.AddRange(theEmail.Body.Split(vbLf)) ' need to know the data, look at a line
        'Dim myArray = ListBox1.Items.OfType(Of String)().ToArray()
        'Dim nnn = Array.IndexOf(myArray, "Salary")
        'vblf was at the beginning of each line using newline or vbcr or whatever, which affected searches
    End Sub

End Class
