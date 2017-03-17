<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.tabControl1 = New System.Windows.Forms.TabControl()
        Me.tabPage2 = New System.Windows.Forms.TabPage()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.listView1 = New System.Windows.Forms.ListView()
        Me.ContactsBox = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.ipBX = New System.Windows.Forms.TextBox()
        Me.FindSBUbtn = New System.Windows.Forms.Button()
        Me.tabPage1 = New System.Windows.Forms.TabPage()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.tabControl1.SuspendLayout()
        Me.tabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.AllowDrop = True
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.Location = New System.Drawing.Point(31, 38)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListBox1.Size = New System.Drawing.Size(610, 524)
        Me.ListBox1.TabIndex = 2
        '
        'tabControl1
        '
        Me.tabControl1.Controls.Add(Me.tabPage2)
        Me.tabControl1.Controls.Add(Me.tabPage1)
        Me.tabControl1.Location = New System.Drawing.Point(675, 38)
        Me.tabControl1.Name = "tabControl1"
        Me.tabControl1.SelectedIndex = 0
        Me.tabControl1.Size = New System.Drawing.Size(646, 557)
        Me.tabControl1.TabIndex = 8
        '
        'tabPage2
        '
        Me.tabPage2.BackColor = System.Drawing.Color.Transparent
        Me.tabPage2.Controls.Add(Me.Label5)
        Me.tabPage2.Controls.Add(Me.listView1)
        Me.tabPage2.Controls.Add(Me.ContactsBox)
        Me.tabPage2.Controls.Add(Me.label2)
        Me.tabPage2.Controls.Add(Me.label1)
        Me.tabPage2.Controls.Add(Me.ipBX)
        Me.tabPage2.Controls.Add(Me.FindSBUbtn)
        Me.tabPage2.Location = New System.Drawing.Point(4, 22)
        Me.tabPage2.Name = "tabPage2"
        Me.tabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPage2.Size = New System.Drawing.Size(638, 531)
        Me.tabPage2.TabIndex = 1
        Me.tabPage2.Text = "IP Lookup"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(24, 121)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 23)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Info:"
        '
        'listView1
        '
        Me.listView1.FullRowSelect = True
        Me.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.listView1.Location = New System.Drawing.Point(23, 147)
        Me.listView1.Name = "listView1"
        Me.listView1.Size = New System.Drawing.Size(562, 347)
        Me.listView1.TabIndex = 13
        Me.listView1.UseCompatibleStateImageBehavior = False
        Me.listView1.View = System.Windows.Forms.View.Details
        '
        'ContactsBox
        '
        Me.ContactsBox.Location = New System.Drawing.Point(142, 90)
        Me.ContactsBox.Name = "ContactsBox"
        Me.ContactsBox.Size = New System.Drawing.Size(444, 20)
        Me.ContactsBox.TabIndex = 12
        Me.ContactsBox.Text = "Contacts.xlsx"
        '
        'label2
        '
        Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.Location = New System.Drawing.Point(24, 87)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(112, 23)
        Me.label2.TabIndex = 11
        Me.label2.Text = "Contacts File:"
        Me.label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'label1
        '
        Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(102, 64)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(34, 23)
        Me.label1.TabIndex = 10
        Me.label1.Text = "IP:"
        Me.label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ipBX
        '
        Me.ipBX.Location = New System.Drawing.Point(142, 64)
        Me.ipBX.Name = "ipBX"
        Me.ipBX.Size = New System.Drawing.Size(444, 20)
        Me.ipBX.TabIndex = 8
        '
        'FindSBUbtn
        '
        Me.FindSBUbtn.Location = New System.Drawing.Point(24, 25)
        Me.FindSBUbtn.Name = "FindSBUbtn"
        Me.FindSBUbtn.Size = New System.Drawing.Size(75, 23)
        Me.FindSBUbtn.TabIndex = 7
        Me.FindSBUbtn.Text = "FindSBU"
        Me.FindSBUbtn.UseVisualStyleBackColor = True
        '
        'tabPage1
        '
        Me.tabPage1.BackColor = System.Drawing.Color.Transparent
        Me.tabPage1.Location = New System.Drawing.Point(4, 22)
        Me.tabPage1.Name = "tabPage1"
        Me.tabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabPage1.Size = New System.Drawing.Size(638, 531)
        Me.tabPage1.TabIndex = 0
        Me.tabPage1.Text = "Instructions"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(31, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(144, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Drag outlook message below"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(308, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'WebBrowser1
        '
        Me.WebBrowser1.AllowWebBrowserDrop = False
        Me.WebBrowser1.Location = New System.Drawing.Point(336, 207)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(427, 347)
        Me.WebBrowser1.TabIndex = 11
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1360, 670)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tabControl1)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "Form1"
        Me.Text = "Cyber Message Processor"
        Me.tabControl1.ResumeLayout(False)
        Me.tabPage2.ResumeLayout(False)
        Me.tabPage2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBox1 As ListBox
    Private WithEvents tabControl1 As TabControl
    Private WithEvents tabPage2 As TabPage
    Private WithEvents listView1 As ListView
    Private WithEvents ContactsBox As TextBox
    Private WithEvents label2 As Label
    Private WithEvents label1 As Label
    Private WithEvents ipBX As TextBox
    Private WithEvents FindSBUbtn As Button
    Private WithEvents tabPage1 As TabPage
    Private WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents WebBrowser1 As WebBrowser
End Class
