Imports System.IO
Public Class CharOption
    Public Char_opt_path As String = Application.StartupPath & "\Char_Options.ini"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim objStreamWriter As StreamWriter
            objStreamWriter = New StreamWriter(Char_opt_path)
            For i As Integer = 0 To ListBox1.Items.Count - 1
                Dim tmp_line As String = ListBox1.Items.Item(i)
                If tmp_line IsNot Nothing Then
                    objStreamWriter.WriteLine(tmp_line)
                End If
            Next
            objStreamWriter.Close()
            MsgBox("Настройки успешно сохранены")
        Catch ex As Exception
            MsgBox("Ошибка! Повторите попытку" & vbCr & ex.Message)
        End Try
    End Sub

    Private Sub CharOption_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AcceptButton = Button1
        Me.CancelButton = Button2
        ListBox1.Items.Clear()
        Dim objStreamReader As StreamReader
        Dim strLine As String
        objStreamReader = New StreamReader(Char_opt_path)
        strLine = objStreamReader.ReadLine
        Do While Not strLine Is Nothing
            ListBox1.Items.Add(strLine)
            strLine = objStreamReader.ReadLine
        Loop
        objStreamReader.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub УдалитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem.Click
        ListBox1.Items.Remove(ListBox1.SelectedItem)
    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click
        With Add_newReplaceChar
            .Text = "Добавить"
            .Find_TB.Text = ""
            .Replace_TB.Text = ""
            .OK_Bt.Text = "Добавить"
            .Cancel_Bt.Text = "Отмена"

            .ShowDialog()
            Dim r As DialogResult = Add_newReplaceChar.DialogResult
            If r = DialogResult.OK Then
                If .Find_TB.Text IsNot "" Then ListBox1.Items.Add(.Find_TB.Text & "=" & .Replace_TB.Text)
            End If
        End With
    End Sub

    Private Sub ИзменитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИзменитьToolStripMenuItem.Click
        LineEdit()
    End Sub
    Sub LineEdit()
        Try
            With Add_newReplaceChar
                .Text = "Изменить"
                .OK_Bt.Text = "Изменить"
                .Cancel_Bt.Text = "Отмена"
                Dim arr() As String = ListBox1.SelectedItem.ToString.Split("=")
                .Find_TB.Text = arr(0).Trim("=")
                .Replace_TB.Text = arr(1).Trim("=")
                .ShowDialog()
                Dim r As DialogResult = Add_newReplaceChar.DialogResult
                If r = DialogResult.OK Then
                    ListBox1.Items.Item(ListBox1.SelectedIndex) = (.Find_TB.Text & "=" & .Replace_TB.Text)
                End If
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        LineEdit()
    End Sub
End Class