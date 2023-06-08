Imports System.Drawing
Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.IO

Public Class Form4

    Public pictureType As String

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.TextBox1.BackColor = System.Drawing.Color.Transparent
        'Me.RichTextBox1.BackColor = System.Drawing.Color.Transparent
        Me.RadioButton1.Select()
        TextBox3.Visible = False
        CheckBox2.Visible = False
        TextBox7.UseSystemPasswordChar = True
        PictureBox1.Image = Global.Excel_VSTO.My.Resources.eye_hide
        pictureType = "hide"
        Dim i As Integer, c As Integer
        c = app.ActiveSheet.Cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        For i = 1 To c
            ComboBox1.Items.Add(app.ActiveSheet.Cells(1, i).Text)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox1.Text = ""
        Me.RichTextBox1.Text = ""
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.ComboBox1.Text = ""
        Me.TextBox6.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            TextBox3.Visible = True
            CheckBox2.Visible = True
        Else
            TextBox3.Visible = False
            CheckBox2.Visible = False
        End If
    End Sub

    Public path As String
    Private Sub TextBox3_DoubleClick(sender As Object, e As EventArgs) Handles TextBox3.DoubleClick
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        If CheckBox2.Checked = True Then
            With FolderBrowserDialog1
                .Description = "请选择附件所在文件夹"
                If .ShowDialog() = vbOK Then
                    TextBox3.Text = .SelectedPath
                    TextBox3.ForeColor = Color.Black
                Else
                    Exit Sub
                End If
            End With
        Else
            With OpenFileDialog1
                .Title = "选择一个或多个文件"
                .Filter = "所有文件|*.*"
                .Multiselect = True
                If .ShowDialog() = vbOK Then
                    TextBox3.Text = String.Join(";", .FileNames)
                    TextBox3.ForeColor = Color.Black
                Else
                    Exit Sub
                End If
            End With
        End If
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        If TextBox3.Text = "请手工输入附件所在目录的完整路径，或双击选择目录" Or TextBox3.Text = "请手工输入文件的完整路径，或双击选择文件" Then
            With TextBox3
                .Text = ""
                .ForeColor = Color.Black
            End With
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myMailsto As New List(Of String), myMailto As String, myMail As String, myPassword As String, mySmtp As String, myPort As String
        Dim mySubject As String, myBody As String, myAttackpath As String, allAttackpath As String
        Dim rng As Excel.Range, rng1 As Excel.Range, coln As Integer, colattack As Integer
        Dim errRecord As New List(Of String)
        Dim result As String

        If TextBox2.Text <> "" And ComboBox2.Text <> "" And TextBox6.Text <> "" And TextBox4.Text <> "" And TextBox7.Text <> "" Then
            myMail = TextBox2.Text & "@" & ComboBox2.Text
            myPassword = TextBox7.Text
            mySmtp = TextBox6.Text
            myPort = TextBox4.Text
            mySubject = TextBox1.Text
            myAttackpath = TextBox3.Text
            If myAttackpath <> "" Then
                RichTextBox1.AppendText(Environment.NewLine)
                RichTextBox1.AppendText(Environment.NewLine)
                RichTextBox1.AppendText("PS:  该邮件应包含附件，如未收到附件请联系发件人")
            End If
            myBody = RichTextBox1.Text
            Select Case CheckBox2.Checked
                Case False
                    If TextBox3.Text = "请手工输入附件所在目录的完整路径，或双击选择目录" Or TextBox3.Text = "请手工输入文件的完整路径，或双击选择文件" Then
                        With TextBox3
                            .Text = ""
                            .ForeColor = Color.Black
                        End With
                    End If

                    '读取收件人地址
                    If TextBox5.Text <> "" Or ComboBox1.Text <> "" Then
                        '读取手工输入收件人地址
                        If TextBox5.Text <> "" Then
                            For Each mail As String In TextBox5.Text.Split(",")
                                myMailsto.Add(mail)
                            Next
                        End If

                        '读取excel表中收件人地址列内容
                        If ComboBox1.Text <> "" Then
                            For Each rng In app.ActiveSheet.range(app.ActiveSheet.cells(1, 1), app.ActiveSheet.cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft))
                                If rng.Value = ComboBox1.Text Then
                                    coln = rng.Column
                                    For Each rng1 In app.ActiveSheet.range(app.ActiveSheet.cells(2, coln), app.ActiveSheet.cells(app.ActiveSheet.rows.count, coln).end(Excel.XlDirection.xlUp))
                                        myMailsto.Add(rng1.Value)
                                    Next
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        MessageBox.Show("收件人邮箱地址不能为空，请核对再次运行")
                        Exit Sub
                    End If

                    '遍历收件人地址，调用发邮件函数发送邮件
                    For Each myMailto In myMailsto
                        result = SendMail(myMailto, myMail, myPassword, mySmtp, myPort, mySubject, myBody, myAttackpath)
                        If result <> "finished" Then
                            errRecord.Add(myMailto & ":" & result)
                        End If
                    Next
                    If errRecord.Count = 0 Then
                        MessageBox.Show("邮件发送成功")
                    Else
                        MessageBox.Show(String.Join(Environment.NewLine, errRecord))
                    End If

                Case True
                    If IO.Directory.Exists(TextBox3.Text) = True Then

                        '按excel表中发件人邮箱地址和附件路径发送邮件
                        If ComboBox1.Text <> "" Then
                            For Each rng In app.ActiveSheet.range(app.ActiveSheet.cells(1, 1), app.ActiveSheet.cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft))
                                If rng.Value = ComboBox1.Text Then
                                    coln = rng.Column
                                End If
                                If rng.Value = "附件" Then
                                    colattack = rng.Column
                                End If
                            Next
                            For Each rng1 In app.ActiveSheet.range(app.ActiveSheet.cells(2, coln), app.ActiveSheet.cells(app.ActiveSheet.rows.count, coln).end(Excel.XlDirection.xlUp))
                                myMailto = rng1.Value
                                allAttackpath = myAttackpath & "\" & app.ActiveSheet.cells(rng1.Row, colattack).value
                                result = SendMail(myMailto, myMail, myPassword, mySmtp, myPort, mySubject, myBody, allAttackpath)
                                If result <> "finished" Then
                                    errRecord.Add(myMailto & ":" & result)
                                End If
                            Next
                            If errRecord.Count = 0 Then
                                MessageBox.Show("邮件发送成功")
                            Else
                                MessageBox.Show(String.Join(Environment.NewLine, errRecord))
                            End If
                        Else
                            MessageBox.Show("收件人邮箱地址不能为空，请核对再次运行")
                            Exit Sub
                        End If
                    Else
                        MessageBox.Show("附件所在文件夹路径错误，请修改后重试")
                    End If
            End Select
        Else
            MessageBox.Show("发件人邮箱不能为空")
        End If
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Select Case ComboBox2.Text
            Case "163.com"
                TextBox6.Text = "smtp.163.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "qq.com"
                TextBox6.Text = "smtp.qq.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "sina.com"
                TextBox6.Text = "smtp.sina.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "sina.cn"
                TextBox6.Text = "smtp.sina.cn"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "126.com"
                TextBox6.Text = "smtp.126.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "sohu.com"
                TextBox6.Text = "smtp.sohu.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "yeah.net"
                TextBox6.Text = "smtp.yeah.net"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "139.com"
                TextBox6.Text = "smtp.139.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "189.cn"
                TextBox6.Text = "smtp.189.cn"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "gmail.com"
                TextBox6.Text = "smtp.gmail.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "outlook.com"
                TextBox6.Text = "smtp-mail.outlook.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "hotmail.com"
                TextBox6.Text = "smtp-mail.outlook.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "aliyun.com"
                TextBox6.Text = "smtp.aliyun.com"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case "wo.cn"
                TextBox6.Text = "smtp.wo.cn"
                TextBox6.ReadOnly = True
                If CheckBox1.Checked = True Then
                    TextBox4.Text = "445"
                    TextBox4.ReadOnly = False
                Else
                    TextBox4.Text = "25"
                    TextBox4.ReadOnly = True
                End If
            Case Else
                TextBox6.Text = ""
                TextBox6.ReadOnly = False
                TextBox4.Text = ""
                TextBox4.ReadOnly = False
        End Select
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            If TextBox6.Text <> "" Then
                TextBox4.Text = "445"
                TextBox4.ReadOnly = False
            Else
                TextBox4.Text = ""
                TextBox4.ReadOnly = False
            End If
        Else
            If TextBox6.Text <> "" Then
                TextBox4.Text = "25"
                TextBox4.ReadOnly = True
            Else
                TextBox4.Text = ""
                TextBox4.ReadOnly = False
            End If
        End If
    End Sub

    Public Function SendMail(ByVal mailTo As String, ByVal mailFrom As String, ByVal password As String, ByVal mailSmtp As String, ByVal smtPort As String, ByVal mailSubject As String, ByVal mailBody As String, Optional ByVal mailAttachPaths As String = "") As String

        Try
            '设置SMTP服务器和发送者信息
            Dim smtpServer As New SmtpClient(mailSmtp)
            Dim mail As New MailMessage With {
                .From = New MailAddress(mailFrom)
            }
            smtpServer.Credentials = New System.Net.NetworkCredential(mailFrom, password)

            '设置收件人和邮件内容
            mail.To.Add(mailTo)
            mail.Subject = mailSubject
            mail.Body = mailBody

            '添加附件
            If mailAttachPaths <> "" Then
                For Each mailAttachPath As String In mailAttachPaths.Split(";")
                    Dim attachment As New Attachment(mailAttachPath)
                    mail.Attachments.Add(attachment)
                Next
            End If

            '发送邮件
            smtpServer.Send(mail)

            '清空收件人和附件列表
            mail.To.Clear()
            mail.Attachments.Clear()

            '关闭SMTP连接
            smtpServer.Dispose()
            Return "finished"
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Select Case pictureType
            Case "hide"
                PictureBox1.Image = Global.Excel_VSTO.My.Resources.eye_open
                TextBox7.UseSystemPasswordChar = False
                pictureType = "open"
            Case "open"
                PictureBox1.Image = Global.Excel_VSTO.My.Resources.eye_hide
                TextBox7.UseSystemPasswordChar = True
                pictureType = "hide"
        End Select
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then



            ToolTip1.SetToolTip(TextBox3, "请手工输入附件所在目录的完整路径，或双击选择目录")
            With TextBox3
                .Text = "请手工输入附件所在目录的完整路径，或双击选择目录"
                .ForeColor = Color.LightGray
            End With
        Else
            ToolTip1.SetToolTip(TextBox3, "请手工输入文件的完整路径，或双击选择文件")
            With TextBox3
                .Text = "请手工输入文件的完整路径，或双击选择文件"
                .ForeColor = Color.LightGray
            End With
        End If
    End Sub

    Private Sub CheckBox2_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckStateChanged
        If CheckBox2.Checked = True Then
            Dim messages As New List(Of String) From {
                "1.发送不同附件只能使用表结构维护,需要在当前excel表中存在‘附件’列。",
                "2.请注意所有附件均应放入同一文件夹，并可使用功能包中的‘批读文件名’读入文件名。",
                "3.将批量读入文件名复制入‘附件’列，并做好与电子邮箱所在列的对应关系",
                "请务必按以上说明操作，否则发送附件将会出错!"
            }
            Dim message As String = String.Join(Environment.NewLine, messages)
            Dim caption As String = "重要提示"
            Dim dr As DialogResult = MessageBox.Show(message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
            If dr = DialogResult.OK Then
                TextBox5.ReadOnly = True
            Else
                CheckBox2.Checked = False
                TextBox5.ReadOnly = False
            End If
        Else
            TextBox5.ReadOnly = False
        End If
    End Sub
End Class