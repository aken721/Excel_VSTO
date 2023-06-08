Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class Form2
    Private Sub TabControl1_DrawItem(sender As Object, e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem

        'Me.TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed ' 必须先改变模式(可将此句放在Load事件上)
        Dim tabArea As System.Drawing.Rectangle
        Dim tabTextArea As RectangleF
        tabArea = TabControl1.GetTabRect(e.Index)
        tabTextArea = CType(TabControl1.GetTabRect(e.Index), RectangleF)

        Dim g As Graphics = e.Graphics
        Dim sf As New StringFormat With {
            .LineAlignment = StringAlignment.Center,
            .Alignment = StringAlignment.Center
        }
        Dim unused As System.Drawing.Font = Me.TabControl1.Font
        Dim font As New System.Drawing.Font("微软雅黑"， 9.0F, FontStyle.Bold)
        Dim brush As New SolidBrush(Color.DarkBlue)
        g.DrawString((CType(sender, TabControl)).TabPages(e.Index).Text, font, brush, tabTextArea, sf)
    End Sub
    Dim thr As Thread
    Private ReadOnly mytime As Long

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'AutoCtlSize(Form2.ActiveForm)
        TabControl1.SelectTab(0)
        Dim wb As Excel.Workbook, ws As Excel.Worksheet
        wb = app.ActiveWorkbook
        ComboBox1.Items.Clear()
        ListBox1.Items.Clear()
        For Each ws In wb.Worksheets
            ComboBox1.Items.Add(ws.Name)
            ListBox1.Items.Add(ws.Name)
        Next
        ComboBox1.Refresh()
        ListBox1.Refresh()
        Label8.Text = ""
        Label8.Visible = False
        Label9.Text = ""
        Label9.Visible = False
        Control.CheckForIllegalCrossThreadCalls = False
    End Sub

    Private Sub ComboBox1_Click(sender As Object, e As EventArgs) Handles ComboBox1.Click
        Label8.Text = ""
        Label8.Visible = False
        Dim wb As Excel.Workbook, ws As Excel.Worksheet
        wb = app.ActiveWorkbook
        ComboBox1.Items.Clear()
        For Each ws In wb.Worksheets
            ComboBox1.Items.Add(ws.Name)
        Next
        ComboBox1.Refresh()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Label8.Text = ""
        Label8.Visible = False
        Dim shtname As String, rng As Excel.Range, ws As Excel.Worksheet
        shtname = ComboBox1.Text
        If shtname = "" Then
            ComboBox2.Items.Clear()
        Else
            ws = app.ActiveWorkbook.Worksheets(shtname)
            If (shtname <> "") Then
                ComboBox2.Items.Clear()
                For Each rng In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Columns.Count).end(Excel.XlDirection.xlToLeft))
                    If Len(rng.Value) > 0 Then
                        ComboBox2.Items.Add(rng.Value)
                    End If
                Next
                If ComboBox2.Items.Count > 0 Then
                    ComboBox2.Text = ComboBox2.Items(0)
                Else
                    ComboBox2.Text = ""
                End If
            Else
                ComboBox2.Text = ""
            End If
        End If
        ComboBox2.Refresh()
    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        Label8.Text = ""
        Label8.Visible = False
        Dim shtname As String, rng As Excel.Range, ws As Excel.Worksheet
        shtname = ComboBox1.Text
        If shtname = "" Then
            ComboBox2.Items.Clear()
        Else
            ws = app.ActiveWorkbook.Worksheets(shtname)
            If (shtname <> "") Then
                ComboBox2.Items.Clear()
                For Each rng In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Columns.Count).end(Excel.XlDirection.xlToLeft))
                    If Len(rng.Value) > 0 Then
                        ComboBox2.Items.Add(rng.Value)
                    End If
                Next
                If ComboBox2.Items.Count > 0 Then
                    ComboBox2.Text = ComboBox2.Items(0)
                Else
                    ComboBox2.Text = ""
                End If
            Else
                ComboBox2.Text = ""
            End If
        End If
        ComboBox2.Refresh()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        Select Case TabControl1.SelectedIndex
            Case 0
                Dim wb As Excel.Workbook, ws As Excel.Worksheet
                Dim combo1text = ComboBox1.Text
                Dim combo2text = ComboBox2.Text
                wb = app.ActiveWorkbook
                ComboBox1.Items.Clear()
                ComboBox2.Items.Clear()
                For Each ws In wb.Worksheets
                    ComboBox1.Items.Add(ws.Name)
                Next
                ComboBox1.Refresh()
                ComboBox2.Refresh()
                ComboBox1.Text = combo1text
                ComboBox2.Text = combo2text
            Case 1

            Case 2
                Dim wb As Excel.Workbook, ws As Excel.Worksheet
                wb = app.ActiveWorkbook
                ListBox1.Items.Clear()
                For Each ws In wb.Worksheets
                    ListBox1.Items.Add(ws.Name)
                Next
                ListBox1.Refresh()
            Case 3

            Case 4

            Case 5
                Me.Dispose()
        End Select

    End Sub

    '分表-根据所选择字段进行分表

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        thr = New Thread(AddressOf Resolve_table)
        thr.Start()
        Label8.Visible = True
        Timer1.Interval = 1000
        Timer1.Enabled = True
    End Sub

    Dim outputpath As String

    '分表中的分表导出为工作簿

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label8.Text = ""
        Label8.Visible = False
        With FolderBrowserDialog1
            .Description = "请选择文件所在文件夹"
            If .ShowDialog() = vbOK Then
                outputpath = .SelectedPath
            Else
                Exit Sub
            End If
        End With
        thr = New Thread(AddressOf Export_table)
        thr.Start()
        Timer2.Interval = 1000
        Timer2.Enabled = True
    End Sub

    '分表中的删除分出的表（同一工作簿下）

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label8.Text = ""
        Label8.Visible = False
        Try
            Dim selesheet As String, selerange As String, selecell As Excel.Range, t As Integer, i As Integer, bname As String, n As Long, coln As Integer, rown As Integer
            Dim arr As New List(Of String)()
            Dim arr1 As New List(Of String)()
            Dim vlu As String, wks As Excel.Worksheet, ws As Excel.Worksheet
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            bname = app.ActiveWorkbook.Name
            selesheet = Me.ComboBox1.Text
            selerange = Me.ComboBox2.Text
            ws = app.Workbooks(bname).Worksheets(selesheet)
            ws.Activate()
            coln = ws.Cells(1, ws.Columns.Count).End(Excel.XlDirection.xlToLeft).column
            ws.Activate()

            '找出分表关键字段在哪一列
            For Each selecell In ws.Range(ws.Cells(1, 1), ws.Cells(1, coln))
                If selecell.Value = selerange Then
                    t = selecell.Column
                    Exit For
                End If
            Next
            rown = ws.Cells(ws.Rows.Count, t).End(Excel.XlDirection.xlUp).Row

            '通过数组去重找出分出的表名
            For i = 1 To rown - 1
                arr.Add(ws.Cells(i + 1, t).Value)
            Next
            For Each keyname As String In arr
                If Not arr1.Contains(keyname) Then
                    arr1.Add(keyname)
                End If
            Next

            '删除分出的表
            For n = 1 To arr1.Count
                vlu = arr1(n - 1)
                For Each wks In app.ActiveWorkbook.Worksheets
                    If wks.Name = vlu Then
                        wks.Delete()
                    End If
                Next
            Next
            app.ScreenUpdating = True
            app.DisplayAlerts = True
        Catch ex As Exception
            MsgBox("删除分表出现错误。错误代码：" & ex.Message)
        End Try
    End Sub

    '分表中的清空选项功能

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Label8.Text = ""
        Label8.Visible = False
        ComboBox1.Items.Clear()
        Dim wb As Excel.Workbook, ws As Excel.Worksheet
        wb = app.ActiveWorkbook
        For Each ws In wb.Worksheets
            ComboBox1.Items.Add(ws.Name)
        Next
        ComboBox1.Refresh()
        Dim shtname As String, rng As Excel.Range
        shtname = ComboBox1.Text
        If shtname = "" Then
            ComboBox2.Items.Clear()
        Else
            ws = app.ActiveWorkbook.Worksheets(shtname)
            If (shtname <> "") Then
                ComboBox2.Items.Clear()
                For Each rng In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Columns.Count).end(Excel.XlDirection.xlToLeft))
                    If Len(rng.Value) > 0 Then
                        ComboBox2.Items.Add(rng.Value)
                    End If
                Next
                If ComboBox2.Items.Count > 0 Then
                    ComboBox2.Text = ComboBox2.Items(0)
                Else
                    ComboBox2.Text = ""
                End If
            Else
                ComboBox2.Text = ""
            End If
        End If
        ComboBox2.Refresh()
    End Sub

    '执行并表-单一工作簿下工作表合并

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Label9.Text = ""
        Label9.Visible = False
        Dim sht As Excel.Worksheet, rng As Excel.Range, rng1 As Excel.Range, totelsheet As Excel.Worksheet
        Dim xrow As Integer, ycolumn As Integer, wbname As String, tt As Integer, s As Integer, t As String
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        wbname = app.ActiveWorkbook.Name
        For Each sht In app.ActiveWorkbook.Worksheets
            If sht.Name = "并表汇总" Then
                MsgBox("已存在名称为'并表汇总'的工作表，请改名或删除该表后再次执行并表程序")
                Exit Sub
            End If
        Next
        totelsheet = app.Workbooks(wbname).Worksheets.Add(Before:=app.Workbooks(wbname).Worksheets(1))
        With totelsheet
            .Name = "并表汇总"
            .Activate()
            .Range("A1").Select()
        End With
        t = InputBox("请输入数据起始行的行数（所输数字应大于等于2,若不输入或输入小于2数字则默认数据起始行为第2行）", "输入数据起始行")
        If IsNumeric(t) = False Then
            s = 0
        Else
            s = Convert.ToInt32(t)
        End If
        If s < 2 Then
            s = 2
        End If
        '在合并表中粘贴标题行
        app.ActiveWorkbook.Worksheets(app.ActiveWorkbook.Worksheets.Count).Rows("1:" & s - 1).Copy(totelsheet.Cells(1, 1))
        '合并各表中数据行
        For Each sht In app.ActiveWorkbook.Worksheets
            If sht.Name <> "并表汇总" Then
                'rng为汇总表A列有数据行下一格
                rng = app.ActiveWorkbook.Worksheets("并表汇总").Range("A" & app.ActiveWorkbook.Worksheets("并表汇总").Rows.Count).End(Excel.XlDirection.xlUp).Offset(1, 0)
                'xrow为所有数据的总行数
                xrow = sht.Cells(sht.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row - s + 1
                'ycolumn为有数据最后一列的标号
                ycolumn = sht.Cells(s - 1, sht.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                '复制sht表中所有有数据的内容到rng
                sht.Range("A" & s).Resize(xrow, ycolumn).Copy(rng)
            End If
        Next

        '重写合并表中的序号列
        app.ActiveWorkbook.Worksheets("并表汇总").Activate
        For Each rng1 In app.ActiveSheet.Range(app.ActiveSheet.Cells(s - 1, 1), app.ActiveSheet.Cells(s - 1, app.ActiveSheet.columns.count).End(Excel.XlDirection.xlToLeft))
            If rng1.Value = "序号" Then
                tt = rng1.Column
                For nb = 1 To app.ActiveSheet.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                    app.ActiveSheet.Cells(nb + 1, tt).Value = nb
                Next
                Exit For
            End If
        Next
        app.ScreenUpdating = True
        app.DisplayAlerts = True
    End Sub

    '选择需合并的工作表所在文件夹
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Label9.Text = ""
        Label9.Visible = False
        With FolderBrowserDialog1
            .Description = "请选择文件所在文件夹"
            If .ShowDialog() = vbOK Then
                Me.TextBox1.Text = .SelectedPath
            End If
        End With
    End Sub

    '执行并表-仅对当前文件夹下各工作簿合并（不含子文件夹）并计时
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        thr = New Thread(AddressOf Merge_workbook1)
        thr.Start()
        Label9.Visible = True
        Timer3.Interval = 1000
        Timer3.Enabled = True
    End Sub

    '执行并表-对各子文件夹下工作簿中表合并并计时
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        thr = New Thread(AddressOf Merge_workbook2)
        thr.Start()
        Label9.Visible = True
        Timer3.Interval = 1000
        Timer3.Enabled = True
    End Sub

    '导出&删除功能中的导出
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim outputpath As String, bname As String
        Dim arr As New List(Of String)
        bname = app.ActiveWorkbook.Name
        With FolderBrowserDialog1
            .Description = "请选择导出到文件夹"
            If .ShowDialog() = vbOK Then
                outputpath = .SelectedPath
            Else
                Exit Sub
            End If
        End With
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        Dim n As Integer, t As Integer, item As String
        t = 0
        For Each item In ListBox1.SelectedItems
            arr.Add(item)
        Next
        For n = 0 To arr.Count - 1
            app.ActiveWorkbook.Worksheets(arr(n)).Copy
            app.ActiveWorkbook.SaveAs(outputpath & "\" & arr(n) & ".xlsx")
            app.ActiveWorkbook.Close()
            app.Workbooks(bname).Activate()
        Next
        app.ScreenUpdating = True
        app.DisplayAlerts = True
        MsgBox("所选分表已导出到指定文件夹", vbOKOnly, "完成！")
    End Sub

    '导出&删除功能中的删除
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim bname As String
        bname = app.ActiveWorkbook.Name
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        Dim arr As New List(Of String)
        Dim n As Integer, t As Integer
        t = ListBox1.SelectedItems.Count
        For Each item In ListBox1.SelectedItems
            arr.Add(item)
        Next
        If t = ListBox1.Items.Count Then
            MsgBox("批量删除时不能一次性删除所有表，需至少保留一张表")
        Else
            For n = 0 To arr.Count - 1
                app.Workbooks(bname).Worksheets(arr(n)).Delete
            Next
            Dim wb As Excel.Workbook, ws As Excel.Worksheet
            wb = app.ActiveWorkbook
            ListBox1.Items.Clear()
            For Each ws In wb.Worksheets
                ListBox1.Items.Add(ws.Name)
            Next
            ListBox1.Refresh()
            app.ScreenUpdating = True
            app.DisplayAlerts = True
        End If
    End Sub

    '导出&删除功能中的全部选中选项

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        If CheckBox2.Checked = True Then
            CheckBox2.Text = "全部取消"
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, True)
            Next
        Else
            CheckBox2.Text = "全部选中"
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, False)
            Next
        End If
    End Sub

    '导出&删除功能中列表选中功能规则
    Private Sub ListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedValueChanged
        If ListBox1.Items.Count <> ListBox1.SelectedItems.Count Then
            If CheckBox2.Text = "全部取消" Then
                CheckBox2.Text = "全部选中"
                CheckBox2.Checked = False
            End If
        Else
            If CheckBox2.Text = "全部选中" Then
                CheckBox2.Text = "全部取消"
                CheckBox2.Checked = True
            End If
        End If
    End Sub

    '同一目录下多工作簿的多工作表汇总到同一工作簿内

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim totelsheet As Excel.Worksheet, AK As Excel.Workbook
        Dim mpath As String, myfile As String, wbname As String, dsheetname As String
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        Me.TopMost = True

        Try
            wbname = app.ActiveWorkbook.Name
            With FolderBrowserDialog1
                .Description = "请选择工作簿所在文件夹"
                If .ShowDialog() = vbOK Then
                    mpath = .SelectedPath
                Else
                    Exit Sub
                End If
            End With

            myfile = Dir(mpath & "\*.xls*")
            Do While myfile <> ""
                If myfile <> app.ActiveWorkbook.Name Then
                    AK = app.Workbooks.Open(mpath & "\" & myfile)
                    For i = 1 To AK.Worksheets.Count
                        If AK.Worksheets(i).UsedRange! = vbEmpty Then
                            dsheetname = Split(AK.Name, ".")(0) & "_" & AK.Worksheets(i).Name
                            AK.Worksheets(i).Rows.Copy
                            app.Workbooks(wbname).Activate()
                            totelsheet = app.Workbooks(wbname).Worksheets.Add(After:=app.Workbooks(wbname).Worksheets(app.Workbooks(wbname).Worksheets.Count))
                            With totelsheet
                                .Name = dsheetname
                                .Range("A1").Select()
                                .Paste()
                            End With
                            totelsheet = Nothing
                            AK.Activate()
                            app.CutCopyMode = False
                        End If
                    Next
                    AK.Close()
                End If
                myfile = Dir()
            Loop
            app.Workbooks(wbname).Activate()
            app.ActiveWorkbook.Sheets(1).Select
            app.ActiveSheet.Range("A1").Select
            app.ActiveWorkbook.Save()
        Catch ex As Exception
            MessageBox.Show("多工作簿表导入出错，错误问题：" & ex.Message)
        Finally
            app.ScreenUpdating = True
            app.DisplayAlerts = True
            Me.TopMost = False
        End Try
    End Sub

    '一键建立指定数量空表

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim totelsheet As Excel.Worksheet, n As Integer, shtname As String, i As Integer
        n = app.InputBox("请输入需要建立空表数量：", "输入建表数量")
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        If n > 0 Then
            shtname = app.InputBox("请输入表统一名称（不能为数字）,未输入则缺省命名为‘新建表’：", "输入表名称")
            If IsNumeric(shtname) Then
                MsgBox("表名输入不合法，将按照缺省名称建表")
                shtname = "新建表"
            ElseIf shtname = "" Then
                shtname = "新建表"
            End If
            For i = 1 To n
                totelsheet = app.ActiveWorkbook.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(app.ActiveWorkbook.Worksheets.Count))
                With totelsheet
                    .Name = shtname & i
                End With
            Next
        End If
        app.ActiveWorkbook.Sheets(1).Activate
        app.ActiveSheet.Range("A1").Select
        app.ActiveWorkbook.Save()
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    '依据指定列进行表转置

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        '多字段转为标识的单一要素记录
        Dim colcount As Integer, rocount As Integer, csheet As String, totalrow As Long, totalcol As Integer, n As Integer, t As Integer, s As Integer, field1 As String, t1 As Integer, transname As String
        Dim transheet As Excel.Worksheet, rng As Excel.Range
        '获取当前表名称
        csheet = app.ActiveSheet.Name
        '获取当前表全部行数
        totalrow = app.ActiveSheet.Rows.Count
        '获取当前表全部列数
        totalcol = app.ActiveSheet.Columns.Count
        '获取最后数据列数
        colcount = app.ActiveSheet.Cells(1, totalcol).End(Excel.XlDirection.xlToLeft).Column
        '获取最后数据行数
        rocount = app.ActiveSheet.Cells(totalrow, 1).End(Excel.XlDirection.xlUp).Row
        transname = csheet & "转置表"
        t = app.InputBox("请输入从第几列（不小于2的数字）开始转置：", "注意")
        If t = False Then
            Exit Sub
        End If
        field1 = app.InputBox("请输入转置列的字段名称：", "注意")
        If field1 = "False" Then
            Exit Sub
        End If
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        For Each rng In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(rocount, colcount))
            If IsNumeric(rng) = True Then
                If rng.Value = "" Or rng.Value = vbNull Then
                    rng.Value = 0
                End If
            End If
        Next
        transheet = app.ActiveWorkbook.Worksheets.Add(Before:=app.ActiveWorkbook.Worksheets(csheet))
        With transheet
            .Name = transname
            For s = 1 To t - 1
                .Cells(1, s) = app.ActiveWorkbook.Worksheets(csheet).Cells(1, s)
            Next
            .Cells(1, t) = "数值"
            .Columns(t).NumberFormatLocal = "#,##0.00"
            .Cells(1, t + 1) = field1
            If field1 = "日期" Then
                .Columns(t + 1).NumberFormatLocal = "yyyy-m-d"
            End If
            .Activate()
        End With
        t1 = t                                                       '使用t1代表每次循环复制转置列的列数
        If t1 = 2 Then
            For n = 1 To colcount - t1 + 1                                '循环重复数据列次
                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Range("A2").Select
                app.ActiveSheet.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown)).Select
                app.Selection.Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, 1).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.ActiveSheet.Paste
                app.CutCopyMode = False

                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Cells(2, t1).Select
                app.Selection.Resize(rocount - 1, 1).Select
                app.Selection.Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, t).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks _
                :=False, Transpose:=False)
                app.CutCopyMode = False

                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Cells(1, t1).Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, t + 1).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.Selection.Resize(rocount - 1, 1).Select
                app.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks _
                :=False, Transpose:=False)
                app.CutCopyMode = False
                t1 += 1
            Next
        Else
            For n = 1 To colcount - t1 + 1                                '循环重复数据列次
                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Range("A2").Select
                app.ActiveSheet.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown)).Select
                app.Selection.Resize(, t - 1).Select
                app.Selection.Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, 1).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.ActiveSheet.Paste
                app.CutCopyMode = False

                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Cells(2, t1).Select
                app.Selection.Resize(rocount - 1, 1).Select
                app.Selection.Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, t).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks _
                :=False, Transpose:=False)
                app.CutCopyMode = False

                app.ActiveWorkbook.Worksheets(csheet).Activate
                app.ActiveSheet.Cells(1, t1).Copy
                app.ActiveWorkbook.Worksheets(transname).Activate
                app.ActiveSheet.Cells(totalrow, t + 1).End(Excel.XlDirection.xlUp).Offset(1, 0).Select
                app.Selection.Resize(rocount - 1, 1).Select
                app.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks _
                :=False, Transpose:=False)
                app.CutCopyMode = False
                t1 += 1
            Next
        End If
        app.ActiveWorkbook.Worksheets(transname).Columns(3).Select
        app.ActiveSheet.Cells(1, 1).Select
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    Public tp As String, pat As String

    '打开正则表达式窗体
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim fm3 As New Form3
        fm3.ShowDialog()
    End Sub

    '生成工资条

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim ws As Excel.Worksheet, acws As String, rown As Long, coln As Integer， wbname As String
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        acws = app.ActiveSheet.Name
        rown = app.ActiveSheet.Range("A1").End(Excel.XlDirection.xlDown).Row
        coln = app.ActiveSheet.Range("A1").End(Excel.XlDirection.xlToRight).Column
        wbname = app.ActiveWorkbook.Name
        app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(rown, coln)).Select
        app.Selection.Copy
        ws = app.Workbooks(wbname).Worksheets.Add(Before:=app.Workbooks(wbname).Worksheets(1))
        With ws
            .Name = "工资条"
            .Activate()
            .Range("A1").PasteSpecial(Excel.XlPasteType.xlPasteAll)
            .Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(rown, coln)).Select()
            app.Selection.EntireColumn.AutoFit
        End With
        For n = rown To 3 Step -1
            app.ActiveSheet.Rows(1).Copy
            app.ActiveSheet.Rows(n).Select
            app.Selection.Insert(Shift:=Excel.XlDirection.xlDown)
            app.CutCopyMode = False
            app.Selection.Insert(Shift:=Excel.XlDirection.xlDown)
        Next
        app.ActiveSheet.Range("A1").Select
        app.ActiveWorkbook.Worksheets(acws).Activate
        app.ActiveSheet.Range("A1").Select
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    '打开工作簿
    Private Sub Load_workbook(loadpath As String, acworkbook As String, sheetname As String)
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        'myfile为路径下excel文件的名称，AK为路径下待合并的workbook，AST为待合并文件中的每一张表，RNG为汇总表中最后一条数据的下一格
        Dim myfile As String, AK As Excel.Workbook, AST As Excel.Worksheet, RNG As Excel.Range, aRow As Integer, i As Integer, n As Integer
        loadpath &= "\"
        myfile = Dir(loadpath & "*.xls*")
        app.ActiveWorkbook.Worksheets(sheetname).Rows(app.ActiveWorkbook.Worksheets(sheetname).Rows.Count).Clear
        n = 1
        Do While myfile <> ""
            If myfile <> app.ActiveWorkbook.Name Then
                AK = app.Workbooks.Open(loadpath & myfile)
                For i = 1 To AK.Sheets.Count
                    AST = AK.Sheets(i)
                    aRow = AST.Range("A" & AST.Rows.Count).End(Excel.XlDirection.xlUp).Row
                    If n = 1 Then
                        RNG = app.Workbooks(acworkbook).Worksheets(sheetname).range("A1")
                        AST.Rows("1:" & aRow).Copy(RNG)
                        n += 1
                    Else
                        RNG = app.Workbooks(acworkbook).Worksheets(sheetname).Range("A" & app.Workbooks(acworkbook).Worksheets(sheetname).rows.count).End(Excel.XlDirection.xlUp).offset(1, 0)
                        AST.Rows("2:" & aRow).Copy(RNG)
                    End If
                Next
                AK.Close(SaveChanges:=False)
            End If
            myfile = Dir()
        Loop
        app.Workbooks(acworkbook).Worksheets(sheetname).Activate
        app.ActiveSheet.Range("A1").Select
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    '判断工作簿是否已退出
    Public Function Existworkbook(wst As Worksheet, ext As String) As Boolean
        Dim rn As Excel.Range
        For Each rn In app.ActiveWorkbook.Worksheets(ext).Range("A1").CurrentRegion
            If wst.Name = rn.Value Then
                Existworkbook = True
                Return Existworkbook
                Exit Function
            End If
        Next
        Existworkbook = False
        Return Existworkbook
    End Function

    '自动调整窗体大小
    Public Sub AutoCtlSize(ByVal inObj As Control)     '自动调整控件大小 
        If inObj Is Nothing Then
            Exit Sub
        End If
        '显示分辨率与窗体工作区的大小的关系：分辨率width*height--工作区（没有工具栏）width*(height-46) 
        '即分辨率为*600时,子窗体只能为*554 
        '上述情况还要windows状态栏自动隐藏，如果不隐藏，则height还要减少，结果为：*524 
        '检测桌面显示分辨率(Visual Basic)请参见 
        '此示例以像素为单位确定桌面的宽度和高度。 
        Dim DeskTopSize As Size = System.Windows.Forms.SystemInformation.PrimaryMonitorSize
        Dim FontSize As Single
        FontSize = 9 * DeskTopSize.Height / 600

        '控件本身**** 
        '控件大小 
        inObj.Size = New Size(Int(inObj.Size.Width * DeskTopSize.Width / 800), Int(inObj.Size.Height * DeskTopSize.Height / 600))
        '控件位置 
        inObj.Location = New System.Drawing.Point(Int(inObj.Location.X * DeskTopSize.Width / 800), Int(inObj.Location.Y * DeskTopSize.Height / 600))
        '如果控件为Form,则设置固定边框 
        'Dim mType As Type 
        'Dim mProperty As System.Reflection.PropertyInfo 
        'mType = inObj.GetType 

        'mProperty = mType.GetProperty("FormBorderStyle") 
        'If Not mProperty Is Nothing Then 
        'MessageBox.Show(mType.ToString) 
        ' mProperty.SetValue(inObj, FormBorderStyle.FixedSingle, Nothing) 
        'End If 
        '子控件===== 
        Dim n As Integer
        For n = 0 To inObj.Controls.Count - 1
            '只调整子控件的字体。（如果调整窗体的字体，再调用窗体的show方法时，会引发resize从而导致控件的大小和布局改变） 
            inObj.Controls.Item(n).Font = New System.Drawing.Font(inObj.Controls.Item(n).Font.FontFamily, FontSize)
            '递归调用（穷举所有子控件） 
            AutoCtlSize(inObj.Controls.Item(n))
        Next
    End Sub


    Dim res As Boolean

    '分表

    Private Sub Resolve_table()
        res = False
        Me.TopMost = True
        Me.Button1.Enabled = False
        Me.Button2.Enabled = False
        Me.Button3.Enabled = False
        Me.Button4.Enabled = False
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Enabled = False
        Me.ControlBox = False

        Try
            Dim selesheet As String, selerange As String, selecell As Excel.Range, i As Integer, bname As String， v As String
            Dim rng As Excel.Range, wks As Excel.Worksheet, nb As Long, tt As Integer
            '声明范围列数、范围行数、分表依据列数、筛选结果第一列数
            Dim coln As Integer, rown As Integer, resolvecoln As Integer
            Dim arr As New List(Of String)()
            Dim arr1 As New List(Of String)()
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            bname = app.ActiveWorkbook.Name
            selesheet = ComboBox1.Text
            selerange = ComboBox2.Text
            app.Workbooks(bname).Worksheets(selesheet).Activate

            ''找出分表关键字段在哪一列
            For Each selecell In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                If selecell.Value = selerange Then
                    resolvecoln = selecell.Column
                    Exit For
                End If
            Next
            coln = app.ActiveSheet.Cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).column
            rown = app.ActiveSheet.Cells(app.ActiveSheet.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row

            '通过数组去重找出应分出的表名
            For i = 1 To rown - 1
                If (Len(app.ActiveSheet.Cells(i + 1, resolvecoln).Value) = 0) Then
                    app.ActiveSheet.Cells(i + 1， resolvecoln) = "空白"
                    arr.Add("空白")
                Else
                    arr.Add(app.ActiveSheet.Cells(i + 1, resolvecoln).Value)
                End If
            Next
            For Each keyname As String In arr
                If Not arr1.Contains(keyname) Then
                    arr1.Add(keyname)
                End If
            Next

            '新建分表，并通过关键字段筛选，筛出结果复制到相应分表中
            For Each v In arr1
                app.ActiveWorkbook.Worksheets.Add(, After:=app.ActiveWorkbook.Worksheets(app.ActiveWorkbook.Worksheets.Count)).Name = v
                app.Workbooks(bname).Worksheets(selesheet).select
                app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, coln)).Select
                app.Selection.AutoFilter(Field:=resolvecoln, Criteria1:=v)
                app.ActiveSheet.Rows(1).Select
                app.ActiveSheet.Range(app.Selection, app.Selection.End(Excel.XlDirection.xlDown)).Select
                app.Selection.Copy(app.ActiveWorkbook.Worksheets(v).Range("A1"))
            Next
            app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, coln)).AutoFilter
            app.ActiveSheet.range("A1").select

            '对有序号列的表中序号重排序
            For Each wks In app.ActiveWorkbook.Worksheets
                wks.Activate()
                For Each rng In wks.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                    If rng.Value = "序号" Then
                        tt = rng.Column
                        For nb = 1 To wks.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                            wks.Cells(nb + 1, tt).Value = nb
                        Next
                        Exit For
                    End If
                Next
            Next
            app.ActiveWorkbook.Worksheets(selesheet).Activate
            app.ActiveSheet.Range("A1").Select
            app.ScreenUpdating = True
            app.CutCopyMode = False
        Catch ex As Exception
            MessageBox.Show("选择的表或字段不正确，请核对后再试。错误问题：" & ex.Message)
        Finally
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
            Me.Button4.Enabled = True
            Me.ComboBox1.Enabled = True
            Me.ComboBox2.Enabled = True
            Me.ControlBox = True
            res = True
        End Try
    End Sub

    '导出表
    Private Sub Export_table()
        res = False
        Me.TopMost = True
        Me.Button1.Enabled = False
        Me.Button2.Enabled = False
        Me.Button3.Enabled = False
        Me.Button4.Enabled = False
        Me.ComboBox1.Enabled = False
        Me.ComboBox2.Enabled = False
        Me.ControlBox = False

        Try
            Dim bname As String, combotextname As String, c1text As String, c2text As String
            c1text = ComboBox1.Text
            c2text = ComboBox2.Text
            bname = app.ActiveWorkbook.Name
            combotextname = ComboBox1.Text
            app.DisplayAlerts = False
            app.ScreenUpdating = False
            Dim sht As Excel.Worksheet
            For Each sht In app.ActiveWorkbook.Worksheets
                If sht.Name <> c1text Then
                    If File.Exists(outputpath & "\" & sht.Name & ".xlsx") Then
                        File.Delete(outputpath & "\" & sht.Name & ".xlsx")
                    End If
                    sht.Copy()
                    app.ActiveWorkbook.SaveAs(outputpath & "\" & sht.Name & ".xlsx")
                    app.ActiveWorkbook.Close()
                    sht.Delete()
                End If
            Next
            Me.TopMost = False
            app.DisplayAlerts = True
            app.ScreenUpdating = True
        Catch ex As Exception
            MessageBox.Show("分表导出出现错误，错误代码：" & ex.Message)
        Finally
            app.CutCopyMode = False
            app.DisplayAlerts = True
            app.ScreenUpdating = True
            Me.Button1.Enabled = True
            Me.Button2.Enabled = True
            Me.Button3.Enabled = True
            Me.Button4.Enabled = True
            Me.ComboBox1.Enabled = True
            Me.ComboBox2.Enabled = True
            Me.ControlBox = True
            res = True
        End Try
    End Sub

    '并表-仅对当前文件夹下各工作簿合并（不含子文件夹）
    Private Sub Merge_workbook1()
        res = False
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        Me.TopMost = True
        Me.Button5.Enabled = False
        Me.Button6.Enabled = False
        Me.Button7.Enabled = False
        Me.Button8.Enabled = False
        Me.TextBox1.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.ControlBox = False

        Try
            Dim mpath As String, totelsheet As Excel.Worksheet, wbname As String, tt As Integer, rng1 As Excel.Range, wks As Excel.Worksheet
            wbname = app.ActiveWorkbook.Name
            mpath = Me.TextBox1.Text

            If Len(mpath) > 0 Then
                If Me.CheckBox1.Checked = False Then
                    totelsheet = app.Workbooks(wbname).Worksheets.Add(Before:=app.Workbooks(wbname).Worksheets(1))
                    With totelsheet
                        .Name = "并表汇总"
                        .Activate()
                        .Range("A1").Select()
                    End With
                    Call Load_workbook(mpath, wbname, "并表汇总")

                    '重新编排序号列
                    app.Worksheets("并表汇总").Activate
                    For Each rng1 In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                        If rng1.Value = "序号" Then
                            tt = rng1.Column
                            For nb = 1 To app.ActiveSheet.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                                app.ActiveSheet.Cells(nb + 1, tt).Value = nb
                            Next
                            Exit For
                        End If
                    Next
                Else
                    totelsheet = app.Workbooks(wbname).Worksheets.Add(Before:=app.Workbooks(wbname).Worksheets(1))
                    With totelsheet
                        .Name = "并表汇总"
                        .Activate()
                        .Range("A1").Select()
                    End With
                    Call Load_workbook(mpath, wbname, "并表汇总")
                    app.Worksheets("并表汇总").Activate

                    For Each wks In app.Workbooks(wbname).Worksheets
                        If wks.Name <> "并表汇总" Then
                            wks.Rows("2:" & wks.Range("A" & wks.Rows.Count).End(Excel.XlDirection.xlUp).Row).copy(app.Worksheets("并表汇总").range("A1").end(Excel.XlDirection.xlDown).offset(1, 0))
                        End If
                    Next
                    '重新编排序号列
                    For Each rng1 In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                        If rng1.Value = "序号" Then
                            tt = rng1.Column
                            For nb = 1 To app.ActiveSheet.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                                app.ActiveSheet.Cells(nb + 1, tt).Value = nb
                            Next
                            Exit For
                        End If
                    Next
                End If
            Else
                MsgBox("未选择文件夹")
            End If
        Catch ex As Exception
            MessageBox.Show("分表合并出现错误，错误代码：" & ex.Message)
        Finally
            app.CutCopyMode = False
            app.DisplayAlerts = True
            app.ScreenUpdating = True
            Me.Button5.Enabled = True
            Me.Button6.Enabled = True
            Me.Button7.Enabled = True
            Me.Button8.Enabled = True
            Me.TextBox1.Enabled = True
            Me.CheckBox1.Enabled = True
            Me.ControlBox = True
            res = True
        End Try
    End Sub

    '并表-对各子文件夹下工作簿中表合并
    Private Sub Merge_workbook2()
        res = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        Me.TopMost = True
        Me.Button5.Enabled = False
        Me.Button6.Enabled = False
        Me.Button7.Enabled = False
        Me.Button8.Enabled = False
        Me.TextBox1.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.ControlBox = False

        Try
            Dim fso, folder, fd As Object, xfile
            Dim mpath As String, totelsheet As Excel.Worksheet, wbname As String, tt As Integer, rng1 As Excel.Range
            Dim i As Integer, sht As Excel.Worksheet, rng As Excel.Range, xrow As Integer, ycolumn As Integer
            Dim arr As New List(Of String)

            If Len(Me.TextBox1.Text) > 0 Then
                'wbname为当前打开workbook的文件名
                wbname = app.ActiveWorkbook.Name
                If Me.CheckBox1.Checked = False Then
                    '该选项下仅对所选目录下的子目录中各表格进行汇总，而不汇总已打开工作簿中的表
                    '将打开的工作簿中已存在表暂时改名
                    For Each sht In app.ActiveWorkbook.Worksheets
                        sht.Name = "old_" & sht.Name
                    Next

                    '建立汇总表
                    totelsheet = app.Workbooks(wbname).Worksheets.Add
                    With totelsheet
                        .Name = "并表汇总"
                        .Activate()
                        .Range("A1").Select()
                    End With
                    mpath = Me.TextBox1.Text
                    fso = CreateObject("scripting.filesystemobject")
                    folder = fso.getfolder(mpath)
                    i = 0

                    '一级目录下文件导入
                    xfile = Dir(mpath & "\*.xls*")
                    app.ActiveWorkbook.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(1)).Name = Split(mpath, "\")(mpath.Split("\").Length - 1)
                    Dim finumber As Integer                                    '不为当前打开Excel文件的数量
                    finumber = 0
                    Do While xfile <> ""
                        finumber += 1
                        xfile = Dir()
                    Loop
                    If finumber > 1 Then
                        Call Load_workbook(mpath, wbname， Split(mpath, "\")(mpath.Split("\").Length - 1))
                    ElseIf finumber = 0 Then
                        app.ActiveWorkbook.Worksheets(Split(mpath, "\")(mpath.Split("\").Length - 1)).Delete()
                    Else
                        If Split(xfile, ".")(0) = wbname Then
                            app.ActiveWorkbook.Worksheets(Split(mpath, "\")(mpath.Split("\").Length - 1)).Delete()
                        Else
                            Call Load_workbook(mpath, wbname， Split(mpath, "\")(mpath.Split("\").Length - 1))
                        End If
                    End If

                    '子目录下文件导入
                    For Each fd In folder.SubFolders
                        app.ActiveWorkbook.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(1)).Name = fd.Name
                        Call Load_workbook(mpath & "\" & fd.Name, wbname, fd.Name)
                    Next

                    '导入的各表进行汇总
                    Dim shtnum As Integer = 1                           '标识，用来判断合并的表是不是第一张
                    For Each sht In app.ActiveWorkbook.Worksheets
                        If sht.Name <> "并表汇总" And Split(sht.Name, "_")(0) <> "old" Then
                            If shtnum = 1 Then
                                shtnum += 1
                                rng = app.ActiveWorkbook.Worksheets("并表汇总").Range("A1")
                                'xrow为A1格后有数据的行数
                                xrow = sht.Range("A1").CurrentRegion.Rows.Count
                                'ycolumn为有数据最后一列的标号
                                ycolumn = sht.Cells(1, sht.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                                '复制sht表中所有有数据的内容到rng
                                sht.Range("A1").Resize(xrow, ycolumn).Copy(rng)
                            Else
                                rng = app.ActiveWorkbook.Worksheets("并表汇总").Range("A" & app.ActiveWorkbook.Worksheets("并表汇总").Rows.Count).End(Excel.XlDirection.xlUp).Offset(1, 0)                         'rng为汇总表A列有数据行下一格
                                'xrow为A1格后有数据的行数-1
                                xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1
                                'ycolumn为有数据最后一列的标号
                                ycolumn = sht.Cells(1, sht.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                                '复制sht表中所有有数据的内容到rng
                                sht.Range("A2").Resize(xrow, ycolumn).Copy(rng)
                            End If
                        End If
                    Next
                    For Each sht In app.ActiveWorkbook.Worksheets
                        If Split(sht.Name, "_")(0) = "old" Then
                            sht.Name = Split(sht.Name, "_")(1)

                            '存入的需汇总表删除，暂不启用
                            'ElseIf sht.Name <> "并表汇总" Then
                            '    sht.Delete()
                        End If
                    Next

                    '汇总表重排序列号
                    app.ActiveWorkbook.Worksheets("并表汇总").Activate
                    For Each rng1 In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                        If rng1.Value = "序号" Then
                            tt = rng1.Column
                            For nb = 1 To app.ActiveSheet.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                                app.ActiveSheet.Cells(nb + 1, tt).Value = nb
                            Next
                            Exit For
                        End If
                    Next
                    app.ActiveWorkbook.Sheets("并表汇总").Move(before:=app.ActiveWorkbook.Sheets(1))

                Else

                    '该选项下对所选目录下的子目录中各表格进行汇总，同时已打开工作簿中的表也将汇总
                    '将打开的工作簿中已存在表格存储在
                    For Each sht In app.ActiveWorkbook.Worksheets
                        arr.Add(sht.Name)
                    Next
                    totelsheet = app.Workbooks(wbname).Worksheets.Add
                    With totelsheet
                        .Name = "并表汇总"
                        .Activate()
                        .Range("A1").Select()
                    End With

                    mpath = Me.TextBox1.Text
                    fso = CreateObject("scripting.filesystemobject")
                    folder = fso.getfolder(mpath)
                    i = 0

                    '一级目录下文件导入
                    xfile = Dir(mpath & "\*.xls*")
                    app.ActiveWorkbook.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(1)).Name = Split(mpath, "\")(mpath.Split("\").Length - 1)
                    Dim finumber As Integer                                    '不为当前打开Excel文件的数量
                    finumber = 0
                    Do While xfile <> ""
                        finumber += 1
                        xfile = Dir()
                    Loop
                    If finumber > 1 Then
                        Call Load_workbook(mpath, wbname， Split(mpath, "\")(mpath.Split("\").Length - 1))
                    ElseIf finumber = 0 Then
                        app.ActiveWorkbook.Worksheets(Split(mpath, "\")(mpath.Split("\").Length - 1)).Delete()
                    Else
                        If Split(xfile, ".")(0) = wbname Then
                            app.ActiveWorkbook.Worksheets(Split(mpath, "\")(mpath.Split("\").Length - 1)).Delete()
                        Else
                            Call Load_workbook(mpath, wbname， Split(mpath, "\")(mpath.Split("\").Length - 1))
                        End If
                    End If


                    '子目录下文件导入
                    For Each fd In folder.SubFolders
                        app.ActiveWorkbook.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(1)).Name = fd.Name
                        Call Load_workbook(mpath & "\" & fd.Name, wbname, fd.Name)
                    Next

                    '导入的各表进行汇总
                    Dim shtnum As Integer = 1                           '标识，用来判断合并的表是不是第一张
                    For Each sht In app.ActiveWorkbook.Worksheets
                        If sht.Name <> "并表汇总" Then
                            If shtnum = 1 Then
                                shtnum += 1
                                rng = app.ActiveWorkbook.Worksheets("并表汇总").Range("A1")
                                'xrow为A1格后有数据的行数
                                xrow = sht.Range("A1").CurrentRegion.Rows.Count
                                'ycolumn为有数据最后一列的标号
                                ycolumn = sht.Cells(1, sht.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                                '复制sht表中所有有数据的内容到rng
                                sht.Range("A1").Resize(xrow, ycolumn).Copy(rng)
                            Else
                                rng = app.ActiveWorkbook.Worksheets("并表汇总").Range("A" & app.ActiveWorkbook.Worksheets("并表汇总").Rows.Count).End(Excel.XlDirection.xlUp).Offset(1, 0)                         'rng为汇总表A列有数据行下一格
                                'xrow为A1格后有数据的行数-1
                                xrow = sht.Range("A1").CurrentRegion.Rows.Count - 1
                                'ycolumn为有数据最后一列的标号
                                ycolumn = sht.Cells(1, sht.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
                                '复制sht表中所有有数据的内容到rng
                                sht.Range("A2").Resize(xrow, ycolumn).Copy(rng)
                            End If
                        End If
                    Next

                    '存入的需汇总表和原有表都删除，暂不启用
                    'For Each sht In app.ActiveWorkbook.Worksheets
                    '    If sht.Name <> "并表汇总" Then
                    '        sht.Delete()
                    '    End If
                    'Next
                    app.ActiveWorkbook.Worksheets("并表汇总").Activate
                    For Each rng1 In app.ActiveSheet.Range(app.ActiveSheet.Cells(1, 1), app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight))
                        If rng1.Value = "序号" Then
                            tt = rng1.Column
                            For nb = 1 To app.ActiveSheet.Rows("1:1").End(Excel.XlDirection.xlDown).Row - 1
                                app.ActiveSheet.Cells(nb + 1, tt).Value = nb
                            Next
                            Exit For
                        End If
                    Next
                    app.ActiveWorkbook.Sheets("并表汇总").Move(before:=app.ActiveWorkbook.Sheets(1))
                End If
            Else
                MsgBox("未选择文件夹")
            End If
        Catch ex As Exception
            MessageBox.Show("分表合并出现错误，错误代码：" & ex.Message)
        Finally
            app.CutCopyMode = False
            app.DisplayAlerts = True
            app.ScreenUpdating = True
            Me.Button5.Enabled = True
            Me.Button6.Enabled = True
            Me.Button7.Enabled = True
            Me.Button8.Enabled = True
            Me.TextBox1.Enabled = True
            Me.CheckBox1.Enabled = True
            Me.ControlBox = True
            res = True
        End Try
    End Sub

    Dim counter As Integer

    '分表命令执行时间

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim min As Integer, sec As Integer, hour As Integer, wks As Excel.Worksheet, exsist_null As Boolean
        If res = True Then
            If Label8.Visible = False Then
                Label8.Visible = True
            End If
            Timer1.Enabled = False
            hour = Int(counter / 3600)
            min = Int(counter / 60)
            sec = counter Mod 60
            For Each wks In app.ActiveWorkbook.Worksheets
                If wks.Name = "空白" Then
                    exsist_null = True
                End If
            Next
            If hour > 0 Then
                If exsist_null = True Then
                    Label8.Text = "分表完成，共用" & hour & "时" & min & "分" & sec & "秒" & Chr(10) & "分表字段有null记录，存入空白命名表"
                Else
                    Label8.Text = "分表完成，共用" & hour & "时" & min & "分" & sec & "秒"
                End If
            ElseIf min > 0 Then
                If exsist_null = True Then
                    Label8.Text = "分表完成，共用" & min & "分" & sec & "秒" & Chr(10) & "分表字段有null记录，存入空白命名表"
                Else
                    Label8.Text = "分表完成，共用" & min & "分" & sec & "秒"
                End If
            Else
                If exsist_null = True Then
                    Label8.Text = "分表完成，共用" & sec & "秒" & Chr(10) & "分表字段有null记录，存入空白命名表"
                Else
                    Label8.Text = "分表完成，共用" & sec & "秒"
                End If

            End If
            If thr.IsAlive = True Then
                thr.Abort()
            End If
            counter = 0
            Me.TopMost = False
        Else
            If Label8.Visible = False Then
                Label8.Visible = True
            End If
            counter += 1
            Dim minutes As Integer = counter \ 60 ' 获取分钟数
            Dim remainingSeconds As Integer = counter Mod 60 ' 获取剩余的秒数
            Dim timeString As String = String.Format("{0}:{1:00}", minutes, remainingSeconds)
            Label8.Text = "请勿退出工具，分表进行中......,已用时" & minutes & "分" & remainingSeconds & "秒"
        End If
    End Sub

    '分表导出命令执行时间
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim min As Integer, sec As Integer, hour As Integer
        If res = True Then
            If Label8.Visible = False Then
                Label8.Visible = True
            End If
            Timer2.Enabled = False
            hour = Int(counter / 3600)
            min = Int(counter / 60)
            sec = counter Mod 60
            If hour > 0 Then
                Label8.Text = "分表导出，共用" & hour & "时" & min & "分" & sec & "秒"
            ElseIf min > 0 Then
                Label8.Text = "分表导出，共用" & min & "分" & sec & "秒"
            Else
                Label8.Text = "分表导出，共用" & sec & "秒"
            End If
            If thr.IsAlive = True Then
                thr.Abort()
            End If
            counter = 0
            Me.TopMost = False
            Me.TopMost = False
        Else
            If Label8.Visible = False Then
                Label8.Visible = True
            End If
            counter += 1
            Label8.Text = "请勿退出工具，分表导出中......,已用时" & counter & "秒"
        End If
    End Sub

    '并表命令执行时间
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Dim min As Integer, sec As Integer, hour As Integer
        If res = True Then
            If Label9.Visible = False Then
                Label9.Visible = True
            End If
            Timer3.Enabled = False
            hour = Int(counter / 3600)
            min = Int(counter / 60)
            sec = counter Mod 60
            If hour > 0 Then
                Label9.Text = "并表完成，共用" & hour & "时" & min & "分" & sec & "秒"
            ElseIf min > 0 Then
                Label9.Text = "并表完成，共用" & min & "分" & sec & "秒"
            Else
                Label9.Text = "并表完成，共用" & sec & "秒"
            End If
            If thr.IsAlive = True Then
                thr.Abort()
            End If
            counter = 0
            Me.TopMost = False
        Else
            If Label9.Visible = False Then
                Label9.Visible = True
            End If
            counter += 1
            Label9.Text = "请勿退出工具，分表合并中......,已用时" & counter & "秒"
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Label8.Text = ""
        Label8.Visible = False
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Label8.Text = ""
        Label8.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Label9.Text = ""
        Label9.Visible = False
    End Sub


    '根据表中目录新建空表
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        Dim wb As Excel.Workbook
        Dim ns As Excel.Worksheet
        Dim ws As Excel.Worksheet
        Dim wsname As String, allrow As Long, lastrow As Long
        wb = app.ActiveWorkbook
        MsgBox("该选项是将只包含一个命名为‘目录’Sheet的excel文件自动生成各页空白表的功能,请确定该文件中只包含一个Sheet且已改名为‘目录’，同时各目录项从第二行开始")
        For Each ws In wb.Worksheets
            If ws.Name = "目录" Then
                allrow = wb.Worksheets("目录").Rows.Count
                lastrow = wb.Worksheets("目录").Cells(wb.Worksheets("目录").Cells(allrow, 1).End(Excel.XlDirection.xlUp).Row, 1).Row
                For i = 2 To lastrow
                    wb.Worksheets("目录").Activate
                    wsname = wb.Worksheets("目录").Cells(i, 1).Value
                    ns = wb.Worksheets.Add(After:=app.ActiveWorkbook.Worksheets(app.ActiveWorkbook.Worksheets.Count))
                    ns.Name = wsname
                    wb.Worksheets("目录").Activate
                Next
                wb.Worksheets("目录").Activate
                With app.ActiveSheet
                    For i = 2 To lastrow
                        wsname = app.ActiveSheet.Cells(i, 1).Value
                        .Hyperlinks.Add(anchor:=app.ActiveSheet.Cells(i, 1), Address:="", SubAddress:=wsname & "!A1", TextToDisplay:=wsname)
                        With app.ActiveSheet.Cells(i, 1)
                            .Font.Name = "微软雅黑"
                            .Font.Size = 12
                            .HorizontalAlignment = XlHAlign.xlHAlignCenter
                            .VerticalAlignment = XlVAlign.xlVAlignCenter
                        End With
                    Next
                End With
                wb.Worksheets("目录").Activate
                With app.Range("A1").Font
                    .Name = "微软雅黑"
                    .Size = 12
                    .Bold = True
                End With
                app.Range("A1").HorizontalAlignment = XlHAlign.xlHAlignCenter
                app.Range("A1").VerticalAlignment = XlVAlign.xlVAlignCenter
                app.DisplayAlerts = True
                app.ScreenUpdating = True
                Exit Sub
            End If
        Next
        app.DisplayAlerts = True
        app.ScreenUpdating = True
        MsgBox("未包含命名为'目录'的表格")
    End Sub

End Class