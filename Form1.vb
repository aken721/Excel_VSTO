Imports System.Diagnostics
Imports System.IO

Public Class Form1

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With FolderBrowserDialog1
            .Description = "请选择MP3文件所在文件夹"
            If .ShowDialog() = vbOK Or vbYes Then
                TextBox1.Text = .SelectedPath
            End If
        End With
        Dim fn
        If (Len(TextBox1.Text) > 0) Then
            fn = Dir(TextBox1.Text & "\*.*")
            If (fn <> "") Then
                Label2.Text = "文件夹中现有格式参考：" & fn
                Label2.Visible = True
            End If
        Else
            Label2.Visible = False
        End If
        fn = vbNull
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (Len(TextBox1.Text) = 0) Then
            MsgBox("未选择文件夹，请核对后再运行"， MsgBoxStyle.OkOnly, Title:="注意！")
            Exit Sub
        Else
            Dim workbook As Excel.Workbook, worksheet As Excel.Worksheet, sht As Excel.Worksheet
            Dim c As Long, fn As String, f1 As String, fs As Integer
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            f1 = TextBox1.Text & "\run.bat"
            If Dir(f1) <> "" Then Kill(f1)
            workbook = app.ActiveWorkbook
            For Each sht In app.ActiveWorkbook.Worksheets
                If sht.Name = "cn" Then
                    sht.Name = "cn_备份"
                End If
            Next
            worksheet = workbook.Worksheets.Add
            With worksheet
                .Name = "cn"
                .Activate()
                .Range("C1").Select()
            End With
            c = 1
            fn = Dir(TextBox1.Text & "\*.*")
            Do While fn <> ""
                worksheet.Range("A" & c).Value = fn
                c += 1
                fn = Dir()
            Loop

            If (RadioButton1.Checked = True) Then
                fs = 0
            ElseIf (RadioButton2.Checked = True) Then
                fs = 1
            Else
                fs = -1
            End If
            Select Case fs
                Case 0
                    Call Resultname1(TextBox1.Text)
                Case 1
                    Call Resultname2(TextBox1.Text)
                Case -1
                    MsgBox("格式未选,请先选文件格式")
            End Select
            Dim i As Integer, code As String
            Dim file As New StreamWriter(f1, False, Encoding.GetEncoding("gb2312"))
            For i = 1 To workbook.ActiveSheet.Cells(workbook.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
                code = worksheet.Range("A" & i).Text
                file.WriteLine(code, Encoding.GetEncoding("gb2312"))
            Next
            file.Close()
            worksheet.Delete()
            For Each sht In workbook.Worksheets
                If sht.Name = "cn_备份" Then
                    workbook.Worksheets("cn_备份").Name = "cn"
                End If
            Next
            app.ScreenUpdating = True
            app.DisplayAlerts = True

            '调用批处理文件改文件名，并且不显示cmd窗口
            Dim startInfo As New ProcessStartInfo With {
                .FileName = f1,
                .UseShellExecute = False,
                .CreateNoWindow = True,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True
            }
            Dim proc As Process = Process.Start(startInfo)
            proc.WaitForExit()
            proc.Close()
            Kill(f1)
            MsgBox("文件名修改完毕")
            System.Diagnostics.Process.Start(TextBox1.Text)
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Visible = False
        RadioButton1.Checked = True
        TextBox1.ReadOnly = True
        Call CheckQuote()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If (Len(TextBox1.Text) = 0) Then
            Label2.Visible = False
        End If
    End Sub


    Public Sub CheckQuote()
        '首先检查ADO有没有安装
        Dim bolFindAdo As Boolean
        For Each refed In app.ActiveWorkbook.VBProject.References
            If refed.Name = "Scripting" Then
                If refed.isbroken Then
                    app.ActiveWorkbook.VBProject.References.Remove(refed) '如引用已损坏，删除
                Else
                    bolFindAdo = True
                    Exit For
                End If
            End If
        Next
        If bolFindAdo = False Then
            '还没安装，现在安装scripting runtime
            app.ActiveWorkbook.VBProject.References.AddFromGuid(Guid:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0)
        End If
    End Sub

    Private Function Resultname1(path1 As String) As Boolean
        Dim a As String, b As Integer, c As String, wb As Excel.Workbook = app.ActiveWorkbook, ws As Excel.Worksheet = app.ActiveSheet, rng As Excel.Range
        Dim path As String, n As Integer
        Dim bb As New Form1, tt

        '判断路径中是否有空格，如有空格，需添加""号
        Dim patharr() As String = Split(path1, "\")
        path = patharr(0)
        For n = 1 To UBound(patharr)
            tt = patharr(n)
            If InStr(tt, " ") > 0 Then
                path = path & "\" & Chr(34) & patharr(n) & Chr(34)
            Else
                path = path & "\" & patharr(n)
            End If
        Next

        For Each rng In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row)
            a = rng.Value
            b = rng.Row
            If InStr(a, "-") = 0 Then
                c = a
            Else
                c = Split(a, "-")(1)
                c = Trim(c)
            End If
            ws.Cells(b, 2) = c
            ws.Cells(b, 3) = "ren " & path & "\" & Chr(34) & ws.Range("A" & b).Text & Chr(34) & Chr(32) & Chr(34) & ws.Range("B" & b).Text & Chr(34)
        Next
        ws.Columns("A:B").delete()
        Return Resultname1 = True
    End Function

    Private Function Resultname2(path1 As String) As Boolean
        Dim a As String, b As Integer, c As String, wb As Excel.Workbook = app.ActiveWorkbook, ws As Excel.Worksheet = app.ActiveSheet, rng As Excel.Range
        Dim path As String, n As Integer
        Dim bb As New Form1, tt

        '判断路径中是否有空格，如有空格，需添加""号
        Dim patharr() As String = Split(path1, "\")
        path = patharr(0)
        For n = 1 To UBound(patharr)
            tt = patharr(n)
            If InStr(tt, " ") > 0 Then
                path = path & "\" & Chr(34) & patharr(n) & Chr(34)
            Else
                path = path & "\" & patharr(n)
            End If
        Next

        For Each rng In ws.Range("A1:A" & ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row)
            a = rng.Value
            b = rng.Row
            If InStr(a, "-") = 0 Then
                ws.Cells(b, 2) = a
            Else
                c = Split(a, "-")(0)
                c = Trim(c)
                ws.Cells(b, 2) = c & ".mp3"
            End If
            ws.Cells(b, 3) = "ren " & path & "\" & Chr(34) & ws.Range("A" & b).Text & Chr(34) & Chr(32) & Chr(34) & ws.Range("B" & b).Text & Chr(34)
        Next
        ws.Columns("A:B").delete()
        Return Resultname2 = True
    End Function
End Class