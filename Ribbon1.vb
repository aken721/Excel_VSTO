Imports Microsoft.Office.Tools.Ribbon
Imports System.Diagnostics
Imports System.IO

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ToggleButton1.Checked = False
        ToggleButton1.Label = "改文件名"
        ToggleButton1.ShowLabel = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim subform1 As New Form1
        subform1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim workbook As Excel.Workbook, worksheet As Excel.Worksheet, sht As Excel.Worksheet
        Dim c As Long, fn As String, f1 As String, path As String
        app.DisplayAlerts = False
        app.ScreenUpdating = False

        If ToggleButton1.Checked = False Then
            With FolderBrowserDialog1
                .Description = "请选择文件所在文件夹"
                If .ShowDialog() = vbOK Then
                    path = .SelectedPath
                Else
                    Exit Sub
                End If
            End With
            If (Len(path) > 0) Then
                f1 = path & "\run.bat"
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
                fn = Dir(path & "\*.*")
                Do While fn <> ""
                    worksheet.Range("A" & c).Value = fn
                    worksheet.Range("B" & c).Value = fn
                    c += 1
                    fn = Dir()
                Loop
                worksheet.Range("A" & c).Value = path
                worksheet.Columns(2).select
            Else
                MsgBox("未选择文件夹")
            End If
            fn = vbNull
        Else
            With FolderBrowserDialog1
                .Description = "请选择需改名文件夹所在上级文件夹"
                If .ShowDialog() = vbOK Then
                    path = .SelectedPath
                Else
                    Exit Sub
                End If
            End With
            If (Len(path) > 0) Then
                f1 = path & "\run.bat"
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
                Dim dirs As New IO.DirectoryInfo(path)
                For Each d As IO.DirectoryInfo In dirs.GetDirectories
                    worksheet.Range("A" & c).Value = d.Name
                    worksheet.Range("B" & c).Value = d.Name
                    c += 1
                Next
                worksheet.Range("A" & c).Value = path
                worksheet.Columns(2).select
            Else
                MsgBox("未选择文件夹")
            End If
            fn = vbNull
        End If
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim b As Integer, wb As Excel.Workbook = app.ActiveWorkbook, ws As Excel.Worksheet = app.ActiveSheet
        Dim path As String, n As Integer, path1 As String
        Dim bb As New Form1, tt
        Dim f1 As String
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        path1 = ws.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Value
        f1 = path1 & "\run.bat"
        If Dir(f1) <> "" Then Kill(f1)
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
        For b = 1 To ws.Range("B" & ws.Rows.Count).End(Excel.XlDirection.xlUp).Row
            ws.Range("C" & b).Value = "ren " & path & "\" & Chr(34) & ws.Range("A" & b).Text & Chr(34) & Chr(32) & Chr(34) & ws.Range("B" & b).Text & Chr(34)
        Next
        ws.Columns("A:B").delete()
        Dim i As Integer, code As String
        Dim file As New StreamWriter(f1, False, Encoding.GetEncoding("gb2312"))

        For i = 1 To ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            code = ws.Range("A" & i).Text
            file.WriteLine(code, Encoding.GetEncoding("gb2312"))
        Next
        file.Close()
        wb.Worksheets("cn").Delete()
        For Each sht In wb.Worksheets
            If sht.Name = "cn_备份" Then
                wb.Worksheets("cn_备份").Name = "cn"
            End If
        Next
        app.ScreenUpdating = True
        app.DisplayAlerts = True

        '调用批处理文件改文件名，并且不显示cmd窗口
        Dim startInfo As New ProcessStartInfo()
        startInfo.FileName = f1
        startInfo.UseShellExecute = False
        startInfo.CreateNoWindow = True
        startInfo.RedirectStandardOutput = True
        startInfo.RedirectStandardError = True
        Dim proc As Process = Process.Start(startInfo)
        proc.WaitForExit()
        proc.Close()
        Kill(f1)
        MsgBox("文件名修改完毕")
        Process.Start(path1)
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim subform2 As New Form2
        subform2.Show()
    End Sub

    Private Sub ToggleButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton1.Click
        If ToggleButton1.Checked = True Then
            ToggleButton1.Image = Global.Excel_VSTO.My.Resources.Resources.Radio_Button_on
            ToggleButton1.Label = "改目录名"
            ToggleButton1.ShowLabel = False
            Label1.Label = "目录名"
        Else
            Me.ToggleButton1.Image = Global.Excel_VSTO.My.Resources.Resources.Radio_Button_off
            ToggleButton1.Label = "改文件名"
            ToggleButton1.ShowLabel = False
            Label1.Label = "文件名"
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs)
        Dim b As Integer, wb As Excel.Workbook = app.ActiveWorkbook, ws As Excel.Worksheet = app.ActiveSheet
        Dim path As String, n As Integer, path1 As String， destpath As String, destpath1 As String
        Dim bb As New Form1, tt
        Dim f1 As String

        With FolderBrowserDialog1
            .Description = "请选择要移动至文件夹"
            If .ShowDialog() = vbOK Then
                destpath1 = .SelectedPath
            Else
                Exit Sub
            End If
        End With

        app.DisplayAlerts = False
        app.ScreenUpdating = False
        path1 = ws.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Value
        f1 = path1 & "\run.bat"
        If Dir(f1) <> "" Then Kill(f1)
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

        Dim patharr1() As String = Split(destpath1, "\")
        destpath = patharr1(0)
        For n = 1 To UBound(patharr1)
            tt = patharr1(n)
            If InStr(tt, " ") > 0 Then
                destpath = path & "\" & Chr(34) & patharr(n) & Chr(34)
            Else
                destpath = path & "\" & patharr(n)
            End If
        Next

        For b = 1 To ws.Range("B" & ws.Rows.Count).End(Excel.XlDirection.xlUp).Row
            ws.Range("C" & b).Value = "move " & path & "\" & Chr(34) & ws.Range("A" & b).Text & Chr(34) & Chr(32) & destpath & "\" & Chr(34) & ws.Range("B" & b).Text & Chr(34)
        Next
        ws.Columns("A:B").delete()
        Dim i As Integer, code As String
        Dim file As New StreamWriter(f1, False, Encoding.GetEncoding("gb2312"))
        Dim proc As System.Diagnostics.Process
        For i = 1 To ws.Cells(ws.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            code = ws.Range("A" & i).Text
            file.WriteLine(code, Encoding.GetEncoding("gb2312"))
        Next
        file.Close()
        wb.Worksheets("cn").Delete()
        For Each sht In wb.Worksheets
            If sht.Name = "cn_备份" Then
                wb.Worksheets("cn_备份").Name = "cn"
            End If
        Next
        app.ScreenUpdating = True
        app.DisplayAlerts = True
        proc = System.Diagnostics.Process.Start(f1)
        proc.WaitForExit()
        proc.Close()
        Kill(f1)
        MsgBox("文件移动完毕")
        System.Diagnostics.Process.Start(destpath)
    End Sub
End Class


