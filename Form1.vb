Imports Outlook = NetOffice.OutlookApi

Public Class Form1
    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.All
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myString As String = TextBox2.Text
        Dim index As Integer = ListBox1.FindString(myString, 1)
        If index <> -1 Then
            ' Select the found item:
            ListBox1.SetSelected(index, True)
            MessageBox.Show(ListBox1.Items(index + 2) & "  : count : " & ListBox1.Items.Count.ToString)
        Else
            MessageBox.Show("Item not found.")
        End If
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
        'vblf was at the beginning of each line using newline or vbcr or whatever, which affected searches
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
