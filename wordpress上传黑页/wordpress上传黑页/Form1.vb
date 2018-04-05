Imports System.ComponentModel
Imports System.Threading


Public Class Form1
    Dim xlapp As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
    Dim xlbook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim a As String = 0
    Dim url As String
    Dim HtmlData As String
    Dim HtmlData2 As Object
    Dim Text1 As String
    Dim Text2 As String
    Dim Text3 As String
    Dim DataTou As String
    Dim DataWei As String
    Dim yemian As Integer = 2
    Dim yeshu As Integer
    Dim pd As Boolean = True
    Dim ses As Integer = 0
    Dim pds As Boolean = True
    Delegate Sub wt(ByVal g As String)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        xlapp.DisplayAlerts = False
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            xlbook = xlapp.Workbooks.Open(OpenFileDialog1.FileName)
            xlsheet = xlbook.Sheets(1)
            For i = 1 To xlsheet.UsedRange.Rows.Count  '表格高度
                ListBox1.Items.Add(xlsheet.Range("A" & i).Value)
                ListBox2.Items.Add(xlsheet.Range("B" & i).Value)
                ListBox3.Items.Add(xlsheet.Range("C" & i).Value)
            Next
        End If
        'TextBox2.Text = xlsheet.UsedRange.Columns.Count 获取表格宽度
        'TextBox3.Text = xlsheet.UsedRange.Rows.Count   获取表格高度

    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Try
            xlapp.ActiveWorkbook.Close(SaveChanges:=True)
        Catch ex As Exception

        End Try
        xlsheet = Nothing
        xlbook = Nothing
        xlapp.Quit()
        xlapp = Nothing
        GC.Collect()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            ListBox4.Items.Clear()
            WebBrowser1.Document.All("user_login").SetAttribute("value", ListBox2.SelectedItem)
            WebBrowser1.Document.All("user_pass").SetAttribute("value", ListBox3.SelectedItem)
            WebBrowser1.Document.Forms(0).InvokeMember("submit")
            While (WebBrowser1.ReadyState = WebBrowserReadyState.Complete)
                Application.DoEvents()
            End While
            WebBrowser1.Navigate(url & "wp-admin/edit.php")
            While (WebBrowser1.ReadyState = WebBrowserReadyState.Complete)
                Application.DoEvents()
            End While
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        WebBrowser1.ScriptErrorsSuppressed = True
        WebBrowser2.ScriptErrorsSuppressed = True
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

        HtmlData = Me.WebBrowser1.DocumentText
        If InStr(HtmlData, "amp;action=edit") > 0 Then
            ye()
            While InStr(HtmlData, "amp;action=edit") > 0
                Text2 = Mid(HtmlData, InStr(HtmlData, "amp;action=edit") - 5, 4)
                Text2 = Replace(Text2, "=", "")
                Text2 = Replace(Text2, Chr(38), "")
                If ListBox4.Items.Contains(Text2) = False Then
                    ListBox4.Items.Add(Text2)
                End If

                HtmlData = Mid(HtmlData, InStr(HtmlData, "&amp;action=edit") + 20, HtmlData.Length)
            End While
            'If yemian < yeshu / 4 Then    修改页面
            If TextBox3.Text = "0" Then
                If yemian < yeshu + 1 Then
                    WebBrowser1.Navigate(url & "wp-admin/edit.php?paged=" & yemian)
                    yemian = yemian + 1
                End If
            Else
                If yemian < TextBox3.Text + 1 Then
                    WebBrowser1.Navigate(url & "wp-admin/edit.php?paged=" & yemian)
                    yemian = yemian + 1
                End If
            End If



        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            WebBrowser2.Navigate(url & "wp-admin/post.php?post=" & ListBox4.Items(ses) & Chr(38) & "action=edit")
            If ProgressBar1.Maximum <> ListBox4.Items.Count Then
                ProgressBar1.Maximum = ListBox4.Items.Count
            End If
            ProgressBar1.Value = ses
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim A As HtmlElement
        For Each A In WebBrowser2.Document.GetElementsByTagName("INPUT")
            If A.DomElement.Type = "submit" And A.DomElement.name = "save" Then
                A.InvokeMember("Click")
            End If
        Next
        If ListBox4.Items.Count > ses Then
            ses = ses + 1
            ListBox4.SelectedIndex = ses
            Button4_Click(Nothing, Nothing)
        Else
            ses = 0
            pds = False
            MsgBox("上传完毕")
            ListBox4.Items.Clear()
            WebBrowser2.Navigate("www.baidu.com")
        End If

    End Sub
    Private Sub ye()
        If pd Then
            Text3 = Mid(WebBrowser1.DocumentText, InStr(WebBrowser1.DocumentText, "last-page"), 100)
            Text1 = Mid(Text3, InStr(Text3, "paged=") + 6, 2)
            If Text1.Length > 1 Then
                If Mid(Text1, 2, 1) = Chr(34) Then
                    Text1 = Mid(Text1, 1, 1)
                    yeshu = CType(Text1, Integer)
                End If
                yeshu = CType(Text1, Integer)
            End If
        End If
        pd = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ListBox4.Items.Clear()
    End Sub

    Private Sub WebBrowser2_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser2.DocumentCompleted
        HtmlData2 = WebBrowser2.DocumentText
        If pds Then
            If InStr(HtmlData2, "action=edit") > 0 Then
                Me.WebBrowser2.Document.GetElementById("content").InnerText = Me.WebBrowser2.Document.GetElementById("content").InnerText & Chr("10") & TextBox1.Text
                Button6_Click(Nothing, Nothing)
            End If
        End If


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        pds = True
        ListBox4.Items.Clear()
        ListBox2.SelectedIndex = ListBox1.SelectedIndex
        ListBox3.SelectedIndex = ListBox1.SelectedIndex
        url = ListBox1.SelectedItem
        WebBrowser1.Navigate(url & "wp-login.php")
    End Sub

    Private Sub ListBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox4.SelectedIndexChanged
        ses = ListBox4.SelectedIndex
        Button4_Click(Nothing, Nothing)
    End Sub
End Class
