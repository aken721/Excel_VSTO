Imports System.Text.RegularExpressions
Imports System.Timers
Imports System.Windows.Forms

Public Class Form3
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        SetWindowPos(hWnd:=1, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
        Label2.Visible = False
        Label3.Visible = False
        TextBox1.Visible = False
        ComboBox1.Text = ""
        Dim i As Integer, c As Integer
        c = app.ActiveSheet.Cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        For i = 1 To c
            ComboBox2.Items.Add(app.ActiveSheet.Cells(1, i).Text)
        Next
        ComboBox2.Text = ""
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ws As Excel.Worksheet, wbname As String, wsname As String, rown As Long, coln As Integer, rng As Excel.Range, col As Integer
        Dim rgx As Regex, fm3 As New Form3
        Dim tp As String, pat As String
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        wbname = app.ActiveWorkbook.Name
        wsname = app.ActiveSheet.name
        ws = app.ActiveSheet
        rown = ws.Range("A" & app.ActiveSheet.Rows.Count).End(Excel.XlDirection.xlUp).Row
        coln = ws.Cells(1, app.ActiveSheet.Columns.Count).End(Excel.XlDirection.xlToLeft).Column

        '窗体内选择需过滤的数据列和过滤规则
        If Me.ComboBox2.Text = "" Then
            col = 0
        Else
            col = Int(Me.ComboBox2.Text)
        End If
        tp = ComboBox1.Text
        pat = ""

        '选择已定义的正则表达式过滤条件，或自行写入过滤规则
        Select Case tp
            Case "数字"
                pat = "\d+\.?\d*"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "英文"
                pat = "[A-Za-z]+"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "中文"
                pat = "[^\x00-\xff]+"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "网址"
                pat = "((http|https):\/\/)?[\w-]+(\.[\w-]+)+([\w.,@?^=%&amp;:/~+#-]*[\w@?^=%&amp;/~+#-])?"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "身份证号"
                pat = "\d{15}$|\d{17}([0-9]|X|x)"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "电子邮箱"
                pat = "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "电话号码"
                pat = "(?:(?:\+|00)86)?1[3-9]\d{9}|(?:0[1-9]\d{1,2}-)?\d{7,8}"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "IP地址"
                pat = "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"
                ws.Range(ws.Cells(1, coln + 1), ws.Cells(rown, coln + 1)).NumberFormatLocal = "@"
            Case "自定义"
                If TextBox1.Text <> "" Then
                    pat = TextBox1.Text
                Else
                    MessageBox.Show("请输入正则表达式过滤规则")
                    Exit Sub
                End If
        End Select

        If TypeName(col) = "Integer" And col < coln + 1 And col > 0 Then
            ws.Range(ws.Cells(1, col), ws.Cells(rown, col)).Select()
            rgx = New Regex(pat)
            Dim matchValue As New List(Of String)
            For Each rng In app.Selection
                matchValue.Clear()
                For Each match As Match In rgx.Matches(rng.Value)
                    matchValue.Add(match.Value)
                Next
                Dim result As String = String.Join("|", matchValue)
                app.Cells(rng.Row, coln + 1) = result
            Next
            ShowLabel(Label3, True, "提取完毕")
            StartTimer()
        Else
            MessageBox.Show("您输入的列数有误，请确认")
        End If
        app.DisplayAlerts = True
        app.ScreenUpdating = True
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "自定义" Then
            Label2.Visible = True
            TextBox1.Visible = True
        Else
            Label2.Visible = False
            TextBox1.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        TextBox1.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub


    Private Sub ComboBox2_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox2.TextUpdate
        If ComboBox2.Text = "" Then
            Exit Sub
        Else
            If IsNumeric(ComboBox2.Text) Then
                If ComboBox2.Text <= ComboBox2.Items.Count And ComboBox2.Text > 0 Then
                    Exit Sub
                Else
                    Do Until ComboBox2.Text = ""
                        MessageBox.Show("你输入的数字超出本表有效列数，请重输")
                        ComboBox2.Text = ""
                    Loop
                End If
            Else
                If ComboBox2.Text <> "" Then
                    MessageBox.Show("你输入的不是数字，请重输")
                    ComboBox2.Text = ""
                End If
            End If
        End If
    End Sub

    '控制label3提示完成提取，并在3秒后隐藏
    Private aTimer As New System.Timers.Timer
    Private Delegate Sub SafeCallDelegate(ByVal label As Label, ByVal Visible As Boolean, ByVal Text As String)

    Private Sub ShowLabel(ByVal label As Label, ByVal Visible As Boolean, ByVal Text As String)
        If label.InvokeRequired Then
            Dim d As New SafeCallDelegate(AddressOf ShowLabel)
            label.Invoke(d, New Object() {label, Visible, Text})
        Else
            label.Visible = Visible
            label.Text = Text
        End If
    End Sub

    Private Sub HideLabel(ByVal label As Label, ByVal Visible As Boolean, ByVal Text As String)
        If label.InvokeRequired Then
            Dim d As New SafeCallDelegate(AddressOf HideLabel)
            label.Invoke(d, New Object() {label, Visible, Text})
        Else
            label.Visible = Visible
            label.Text = Text
        End If
    End Sub

    Private Sub StartTimer()
        aTimer.Interval = 3000 '5 seconds
        AddHandler aTimer.Elapsed, AddressOf OnTimedEvent
        aTimer.Enabled = True
    End Sub

    Private Sub OnTimedEvent(source As Object, e As ElapsedEventArgs)
        HideLabel(Label3, False, "")
    End Sub
End Class