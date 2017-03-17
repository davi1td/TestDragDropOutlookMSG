Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports NetOffice.OutlookApi.Enums
Imports Outlook = NetOffice.OutlookApi
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports System.Text
Imports AngleSharp
Imports AngleSharp.Parser.Html
Imports System.IO
Imports AngleSharp.Dom
Imports AngleSharp.Dom.Html

Public Class Form1
    'Dim MailContactInfo As String
    Private MailContactInfoStruc As MailInfoStruc
    Private SymFindDataStruc As SymcEndPFindingsStruc
    Const testing As Boolean = True
    Public testingIP As String = "192.168.1.200"
    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.All
    End Sub
    Structure SymcEndPFindingsStruc
        Public IP As List(Of String)
        Public Host As List(Of String)
        Public Domain As List(Of String)
        Public Users As List(Of String)
    End Structure
    Structure MailInfoStruc
        Public primeFullName As String
        Public primeFirstName As String
        Public primeEmail As String
        Public ccFullName As String
        Public ccEmail As String
        Public SBU As String
        Public location As String
    End Structure
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FirstExample()
    End Sub
    Private Sub FirstExample()
        Dim config = New Configuration().WithCss()
        Dim options As New HtmlParserOptions


        Dim parser = New HtmlParser(config)
        '.Parse().QuerySelectorAll("img").Css("width", "100%")

        Dim myDoc
        Using reader As StreamReader = File.OpenText("F:\MyVisualStudioProjects\TestDragDropOutlookMSG\test.html")
            myDoc = parser.Parse(reader.ReadToEnd)
        End Using
        Dim mydata = myDoc.GetElementsByClassName("mep")
        Dim theIPs = myDoc.GetElementsByClassName("ip_blue") '(0).textcontent   

        SymFindDataStruc.IP = New List(Of String)()
        SymFindDataStruc.Host = New List(Of String)()
        SymFindDataStruc.Domain = New List(Of String)()
        SymFindDataStruc.Users = New List(Of String)()
        For Each ip In theIPs
            If SymFindDataStruc.IP.Find(Function(x) x.Equals(ip.textcontent.Trim)) Is Nothing Then
                SymFindDataStruc.IP.Add(ip.textcontent.Trim)
            End If
        Next
        For Each myDataFields In mydata
            For Each datarows In myDataFields.children
                If (TypeName(datarows) = "HtmlTableElement" AndAlso datarows.rows().length > 1) Then
                    For i = 0 To datarows.rows(0).cells.length - 1

                        Select Case datarows.rows(0).cells(i).innerhtml.trim()
                            Case "Host Name"
                                If SymFindDataStruc.Host.Find(Function(x) x.Equals(datarows.rows(1).cells(i).innerhtml.trim())) Is Nothing Then
                                    SymFindDataStruc.Host.Add(datarows.rows(1).cells(i).innerhtml.trim())
                                End If
                            Case "Domain"
                                If SymFindDataStruc.Domain.Find(Function(x) x.Equals(datarows.rows(1).cells(i).innerhtml.trim())) Is Nothing Then
                                    SymFindDataStruc.Domain.Add(datarows.rows(1).cells(i).innerhtml.trim())
                                End If
                            Case "User(s)"
                                If SymFindDataStruc.Users.Find(Function(x) x.Equals(datarows.rows(1).cells(i).innerhtml.trim())) Is Nothing Then
                                    SymFindDataStruc.Users.Add(datarows.rows(1).cells(i).innerhtml.trim())
                                End If

                        End Select
                    Next


                End If
            Next
        Next

        Dim document = parser.Parse("<h1>Some example source</h1><p>This is a paragraph element")

        'Do something with document like the following

        Console.WriteLine("Serializing the (original) document:")
        Console.WriteLine(document.DocumentElement.OuterHtml)
        'sheet.innerHTML = "div {border: 2px solid black; background-color: blue;}"
        Dim p = document.CreateElement("p")
        Dim p2 = document.CreateElement("p")
        p.Id = "testy"
        '{border: 2px solid black; background-color: blue;}")

        p.TextContent = "This Is first text."
        p.SetAttribute("style", "background-color:powderblue;")
        'p.Style = "color: red;"
        Console.WriteLine("Inserting another element In the body ...")
        document.Body.AppendChild(p)

        Dim oNewP = document.CreateElement("p")
        oNewP.Id = ("whatever")
        Dim oText = document.CreateTextNode("www.java2s.com")
        oNewP.AppendChild(oText)

        Dim beforeMe = document.GetElementsByTagName("p")(0)
        document.Body.InsertBefore(oNewP, beforeMe)


        Console.WriteLine("Serializing the document again:")
        Console.WriteLine(document.DocumentElement.OuterHtml)


        ' WebBrowser1.DocumentText = document.DocumentElement.OuterHtml 'document.Body.OuterHtml
    End Sub
    Private Sub ProcessListBox()

        Dim count As Integer = (ListBox1.Items.Count - 1)
        Dim words As String
        'Dim IP, Host, Domain, Users As New List(Of String)()

        SymFindDataStruc.IP = New List(Of String)()
        SymFindDataStruc.Host = New List(Of String)()
        SymFindDataStruc.Domain = New List(Of String)()
        SymFindDataStruc.Users = New List(Of String)()

        ListBox1.SelectedItem = Nothing
        For a = 0 To count
            words = ListBox1.Items.Item(a)
            If InStr(words.ToLower, "Source IP:".ToLower) Then
                Dim match = Regex.Match(words, "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")
                If (match.Success) Then
                    If SymFindDataStruc.IP.Find(Function(x) x.Equals(match.Value)) Is Nothing Then
                        SymFindDataStruc.IP.Add(match.Value.Trim)

                    End If
                End If
                Console.WriteLine(match.Value)
                'Host Name	User(s)
            ElseIf InStr(words.ToLower, "Host Name	 Domain 	User(s)".ToLower) Then
                'Console.WriteLine(ListBox1.Items.Item(a + 1))
                Dim theString = ListBox1.Items.Item(a + 1).ToString.Trim
                If theString.Split(vbTab).Count > 0 Then
                    If SymFindDataStruc.Host.Find(Function(x) x.Equals(theString.Split(vbTab)(0).Trim)) Is Nothing Then
                        SymFindDataStruc.Host.Add(theString.Split(vbTab)(0).Trim)
                    End If
                    If theString.Split(vbTab).Count > 1 Then
                        If SymFindDataStruc.Domain.Find(Function(x) x.Equals(theString.Split(vbTab)(1).Trim)) Is Nothing Then
                            SymFindDataStruc.Domain.Add(theString.Split(vbTab)(1).Trim)
                        End If
                    End If
                    If theString.Split(vbTab).Count > 2 Then
                        If SymFindDataStruc.Users.Find(Function(x) x.Equals(theString.Split(vbTab)(2).Trim)) Is Nothing Then
                            SymFindDataStruc.Users.Add(theString.Split(vbTab)(2).Trim)
                        End If
                    End If
                End If
            ElseIf InStr(words.ToLower, "host name	 user(s)".ToLower) Then
                'Console.WriteLine(ListBox1.Items.Item(a + 1))
                Dim theString = ListBox1.Items.Item(a + 1).ToString.Trim
                If theString.Split(vbTab).Count > 0 Then
                    If SymFindDataStruc.Host.Find(Function(x) x.Equals(theString.Split(vbTab)(0).Trim)) Is Nothing Then
                        SymFindDataStruc.Host.Add(theString.Split(vbTab)(0).Trim)
                    End If
                    If theString.Split(vbTab).Count > 1 Then
                        If SymFindDataStruc.Users.Find(Function(x) x.Equals(theString.Split(vbTab)(1).Trim)) Is Nothing Then
                            SymFindDataStruc.Users.Add(theString.Split(vbTab)(1).Trim)
                        End If
                    End If

                End If
            End If
        Next
        'Users.Find(Function(x) x.Equals ("SYSTEM"))
        ' send a special warning regarding system account
        If SymFindDataStruc.IP.Count > 1 Then
            listView1.Items.Add("Warning more than one IP encountered, please review further before sending, IPs: " & String.Join(", ", SymFindDataStruc.IP)).ForeColor = Color.Red
            'msgbox warning we're using first IP but found, more than 1 encountered please review before sending
            'show list of IP's
        End If
        If SymFindDataStruc.IP.Count > 0 Then ipBX.Text = SymFindDataStruc.IP(0)
        FindIPInfo(testingIP)
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
        SymFindDataStruc.IP = New List(Of String)()
        SymFindDataStruc.Host = New List(Of String)()
        SymFindDataStruc.Domain = New List(Of String)()
        SymFindDataStruc.Users = New List(Of String)()
        Dim app As Outlook.Application = New Outlook.Application()
        'Dim myNamespace As Outlook.NameSpace = app.GetNamespace("MAPI")
        'Dim test = myNamespace.GetDefaultFolder(OlDefaultFolders.olFolderDrafts)
        Dim config = New Configuration().WithCss()
        'Dim options As New HtmlParserOptions
        Dim parser = New HtmlParser(config)
        '.Parse().QuerySelectorAll("img").Css("width", "100%")

        Dim myMailItem As Outlook.MailItem = app.CreateItem(OlItemType.olMailItem)
        myMailItem.BodyFormat = OlBodyFormat.olFormatHTML
        myMailItem.Attachments.Add(theEmail)
        myMailItem.To = "davi2td@gmail.com"
        myMailItem.Subject = theEmail.Subject
        ' myMailItem.SendUsingAccount("davi2td@gmail.com") 'see https://msdn.microsoft.com/en-us/library/office/ff869311.aspx
        ' Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument
        'Dim htmlStr As New StringBuilder


        Dim testring As String = "<head><title>hi</title><body><table width='200' border='0'><tr><td bgcolor='#999999'>test</td></tr><tr><td bgcolor='#CC99FF'>test</td></tr></table></body></html>"
        Dim myDoc = parser.Parse(theEmail.HTMLBody)
        Dim myDoc2 = parser.Parse("")

        Dim mydata = myDoc.GetElementsByClassName("mep")
        Dim theIPs = myDoc.GetElementsByClassName("ip_blue") '(0).textcontent 
        For Each ip In theIPs
            If SymFindDataStruc.IP.Find(Function(x) x.Equals(ip.TextContent.Trim)) Is Nothing Then
                SymFindDataStruc.IP.Add(ip.TextContent.Trim)
            End If
        Next
        For Each myDataFields In mydata
            For Each datatype In myDataFields.Children

                If (TypeName(datatype) = "HtmlTableElement") Then
                    Dim datarows As IHtmlTableElement = DirectCast(datatype, IHtmlTableElement)
                    If datarows.Rows().Length > 1 Then
                        For i = 0 To datarows.Rows(0).Cells.Length - 1

                            Select Case datarows.Rows(0).Cells(i).InnerHtml.Trim()
                                Case "Host Name"
                                    If SymFindDataStruc.Host.Find(Function(x) x.Equals(datarows.Rows(1).Cells(i).InnerHtml.Trim())) Is Nothing Then
                                        SymFindDataStruc.Host.Add(datarows.Rows(1).Cells(i).InnerHtml.Trim())
                                    End If
                                Case "Domain"
                                    If SymFindDataStruc.Domain.Find(Function(x) x.Equals(datarows.Rows(1).Cells(i).InnerHtml.Trim())) Is Nothing Then
                                        SymFindDataStruc.Domain.Add(datarows.Rows(1).Cells(i).InnerHtml.Trim())
                                    End If
                                Case "User(s)"
                                    If SymFindDataStruc.Users.Find(Function(x) x.Equals(datarows.Rows(1).Cells(i).InnerHtml.Trim())) Is Nothing Then
                                        SymFindDataStruc.Users.Add(datarows.Rows(1).Cells(i).InnerHtml.Trim())
                                    End If

                            End Select
                        Next
                    End If
                End If
            Next
        Next
        Dim test As Object
        If SymFindDataStruc.IP.Count > 0 Then
            test = FindIPAndcontactInfo(testingIP) 'SymFindDataStruc.IP(0))
        End If
        Dim p = myDoc2.CreateElement("p")
        Dim p2 = myDoc2.CreateElement("p")
        p.Id = "testy"

        '{border: 2px solid black; background-color: blue;}")

        p.TextContent = "This Is first text."
        p.SetAttribute("style", "background-color:powderblue;")
        'p.Style = "color: red;"
        Console.WriteLine("Inserting another element In the body ...")
        myDoc2.Body.AppendChild(p)
        p.TextContent = "something else!."
        myDoc2.Body.AppendChild(p)
        Dim oNewP = myDoc2.CreateElement("p")
        oNewP.Id = ("whatever")
        Dim oText = myDoc2.CreateTextNode("www.java2s.com<br>")
        myDoc2.Body.AppendChild(oText)
        oText = myDoc2.CreateTextNode("sdfsdfdfsd")
        myDoc2.Body.AppendChild(oText)

        'Dim beforeMe = myDoc.GetElementsByTagName("p")(0)
        'myDoc.Body.AppendChild(oNewP)
        'listView1.Items.Add(String.Format("Host(s): {0}", String.Join(", ", SymFindDataStruc.Host)))
        'listView1.Items.Add(String.Format("Ip(s): {0}", String.Join(", ", SymFindDataStruc.IP)))
        'listView1.Items.Add(String.Format("User(s) {0}", String.Join(", ", SymFindDataStruc.Users)))
        'If InStr(String.Join(", ", SymFindDataStruc.Users).ToLower, "system") > 0 Then
        '    listView1.Items.Add("The associated user is 'system' which would indicate that the possible infection has elevated privileges.").ForeColor = Color.Red
        '    'ObjectListView1.Items.Add("The associated user Is 'system' which would indicate that the possible infection has elevated privileges.").ForeColor = Color.Red
        'End If
        'listView1.Items.Add("Please let us know what findings are gathered and corrective action taken.")
        'listView1.Items.Add("Regards,")

        myMailItem.HTMLBody = myDoc2.DocumentElement.OuterHtml
        myMailItem.Save()
        Exit Sub


        ListBox1.Items.Clear()
        listView1.Clear()
        ipBX.Text = ""
        'If theEmail .Subject .      "Symantec Endpoint Reported Infection"
        'If SymFindDataStruc.Users.Find(Function(x) x.Equals(ListBox1.Items.Item(a + 1).ToString.Split(vbTab)(2))) Is Nothing Then
        If Not (InStr(theEmail.Subject.ToLower, "Symantec Endpoint Reported Infection".ToLower) > 0) Then
            listView1.Items.Add("'Symantec Endpoint Reported Infection' not found in subject, cannot continue").ForeColor = Color.Red
            Exit Sub
        End If
        ListBox1.Items.AddRange(theEmail.Body.Split(vbLf)) ' need to know the data, look at a line
        'vblf was at the beginning of each line using newline or vbcr or whatever, which affected searches
        ProcessListBox()

    End Sub

    Private Function FindIPAndcontactInfo(ipAddress As String) As MailInfoStruc
        If Not testing And ipAddress = "" Then
            listView1.Items.Add("No IP to lookup found").ForeColor = Color.Red
            Return Nothing
        End If
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False
        Dim utils As Excel.Tools.CommonUtils = New Excel.Tools.CommonUtils(excelApplication)
        Try
            ' add a new workbook
            Dim contactFile As String = My.Application.Info.DirectoryPath + "\" + Path.GetFileName("Contacts.xlsx")
            Dim workBook As Excel.Workbook = excelApplication.Workbooks.Open(contactFile, vbNull, True)
            Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)
            Dim ContactsSheet As Excel.Worksheet = workBook.Worksheets(2)
            Dim ipfuncs As New ipFunctions
            'Dim IptoFind As String = ipBX.Text
            If testing Then ipAddress = testingIP

            For Each cell As Excel.Range In workSheet.UsedRange.Columns(1).Cells
                'Debug.Print(cell.Value+" "+cell.Offset(1,3).Value)
                If ipfuncs.IpIsInSubnet(ipAddress, cell.Value + " " + cell.Offset(1, 3).Value) Then

                    Return FindContactInfo(ContactsSheet, cell.Offset(1, 5).Value)
                End If
            Next
        Catch
            Return Nothing
        End Try
        Return Nothing
    End Function

    Private Function FindIPInfo(ipAddress As String) As Boolean
        listView1.Items.Clear()
        If Not testing And ipAddress = "" Then
            listView1.Items.Add("No IP to lookup found").ForeColor = Color.Red
            Return False
        End If
        Dim excelApplication As New Excel.Application()
        Try
            excelApplication.DisplayAlerts = False
            ' create a utils instance, not need for but helpful to keep the lines of code low
            Dim utils As Excel.Tools.CommonUtils = New Excel.Tools.CommonUtils(excelApplication)

            ' add a new workbook
            Dim contactFile As String = My.Application.Info.DirectoryPath + "\" + Path.GetFileName(ContactsBox.Text)
            Dim workBook As Excel.Workbook = excelApplication.Workbooks.Open(contactFile, vbNull, True)
            Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)
            Dim ContactsSheet As Excel.Worksheet = workBook.Worksheets(2)
            Dim ipfuncs As New ipFunctions
            Dim IptoFind As String = ipBX.Text
            If testing Then IptoFind = testingIP
            For Each cell As Excel.Range In workSheet.UsedRange.Columns(1).Cells
                'Debug.Print(cell.Value+" "+cell.Offset(1,3).Value)
                If ipfuncs.IpIsInSubnet(IptoFind, cell.Value + " " + cell.Offset(1, 3).Value) Then
                    listView1.Items.Add("Found: " + IptoFind + " on row#: " + cell.Row.ToString).ForeColor = Color.Green
                    listView1.Items.Add("SBU = " + cell.Offset(1, 2).Value)
                    listView1.Items.Add("---- Contact Info: ---")
                    MailContactInfoStruc = FindContactInfo(ContactsSheet, cell.Offset(1, 5).Value)
                    listView1.Items.Add(MailContactInfoStruc.primeFirstName)
                    listView1.Items.Add(MailContactInfoStruc.primeEmail)
                    listView1.Items.Add(MailContactInfoStruc.ccEmail)
                    listView1.Items.Add(MailContactInfoStruc.SBU)
                    listView1.Items.Add(MailContactInfoStruc.location)
                    listView1.Items.Add("---- Threat Info: ---")
                    listView1.Items.Add(String.Format("Host(s): {0}", String.Join(", ", SymFindDataStruc.Host)))
                    listView1.Items.Add(String.Format("Ip(s): {0}", String.Join(", ", SymFindDataStruc.IP)))
                    listView1.Items.Add(String.Format("User(s) {0}", String.Join(", ", SymFindDataStruc.Users)))
                    listView1.Items.Add("---- Email to send data: ---")
                    listView1.Items.Add("")
                    listView1.Items.Add("sendTo:" & vbTab & MailContactInfoStruc.primeEmail)
                    listView1.Items.Add("CCTo:" & vbTab & MailContactInfoStruc.ccEmail)
                    listView1.Items.Add(MailContactInfoStruc.primeFirstName & ",")
                    listView1.Items.Add("This is to inform you of the below infection alerts from Symantec on:")
                    listView1.Items.Add(String.Format("Host(s): {0}", String.Join(", ", SymFindDataStruc.Host)))
                    listView1.Items.Add(String.Format("Ip(s): {0}", String.Join(", ", SymFindDataStruc.IP)))
                    listView1.Items.Add(String.Format("User(s) {0}", String.Join(", ", SymFindDataStruc.Users)))
                    If InStr(String.Join(", ", SymFindDataStruc.Users).ToLower, "system") > 0 Then
                        listView1.Items.Add("The associated user is 'system' which would indicate that the possible infection has elevated privileges.").ForeColor = Color.Red
                        'ObjectListView1.Items.Add("The associated user Is 'system' which would indicate that the possible infection has elevated privileges.").ForeColor = Color.Red
                    End If
                    listView1.Items.Add("Please let us know what findings are gathered and corrective action taken.")
                    listView1.Items.Add("Regards,")
                    listView1.Scrollable = True
                    listView1.View = View.Details
                    Dim header = New ColumnHeader()
                    header.Width = listView1.Size.Width - 20
                    header.Text = ""
                    header.Name = "col1"
                    listView1.Columns.Add(header)

                    Return True
                End If
            Next

            listView1.Items.Add("IP Not found in: " & ContactsBox.Text).ForeColor = Color.Red

        Catch ex As Exception
            Dim li As New ListViewItem
            li.Text = "Error: " & ex.ToString
            li.SubItems(0).Font = New Font(li.SubItems(0).Font, li.SubItems(0).Font.Style Or FontStyle.Bold)
            li.ForeColor = Color.Red
            listView1.Items.Add(li)
        Finally
            'cleanup excel stuff
            excelApplication.Quit()
            excelApplication.Dispose()
        End Try
    End Function
    Function FindContactInfo(theSheet As Excel.Worksheet, IndexStr As String) As MailInfoStruc
        Dim MailInfoStruc2 As New MailInfoStruc
        If IndexStr = "" Then Return Nothing
        Dim FoundRow = theSheet.Columns(1).Find((IndexStr.ToUpper))
        If FoundRow Is Nothing Then Return Nothing
        If Not FoundRow Is Nothing Then
            'Return FoundRow.Cells.Offset(1,4).Value + " "+ FoundRow.Cells.Offset(1,5).Value + ";" + FoundRow.Cells.Offset(1,6).Value + " " + FoundRow.Cells.Offset(1,7).Value
            MailInfoStruc2.SBU = FoundRow.Cells.Offset(1, 2).Value
            MailInfoStruc2.location = FoundRow.Cells.Offset(1, 3).Value
            MailInfoStruc2.primeFullName = FoundRow.Cells.Offset(1, 4).Value
            MailInfoStruc2.primeFirstName = MailInfoStruc2.primeFullName.Split(" ")(0)
            MailInfoStruc2.primeEmail = FoundRow.Cells.Offset(1, 5).Value
            MailInfoStruc2.ccFullName = FoundRow.Cells.Offset(1, 6).Value
            MailInfoStruc2.ccEmail = FoundRow.Cells.Offset(1, 7).Value
            Return MailInfoStruc2
        End If
        Return Nothing
    End Function


    Private Sub FindSBUbtn_Click(sender As Object, e As EventArgs) Handles FindSBUbtn.Click
        'FindIPInfo()
    End Sub

    Private Sub AddContextMenu()

        Dim Contextmenu1 As New ContextMenu

        Dim menuItem2Copy As New MenuItem("Copy")
        AddHandler menuItem2Copy.Click, AddressOf menuItem2Copy_Click

        Contextmenu1.MenuItems.Add(menuItem2Copy)

        ' RichTextBox1.ContextMenu = Contextmenu1
        'listView1.ContextMenu = Contextmenu1


    End Sub
    Private Sub menuItem2Copy_Click()
        ' RichTextBox1.Copy()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        AddContextMenu()
    End Sub


End Class
