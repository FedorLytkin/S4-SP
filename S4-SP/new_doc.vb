Imports System.Threading
Public Class new_doc
    Public Arch_ID As Integer = get_reesrt_value("Arch_ID", "7")
    Public ExcelFullFileName As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        InitializeOpenFileDialog()
    End Sub
    Public Sub InitializeOpenFileDialog()
        ' Set the file dialog to filter for graphics files.
        Me.OpenFileDialog1.Filter = "Таблицы(*.xlsx)|"


        ' Allow the user to select multiple images.
        Me.OpenFileDialog1.Multiselect = False
        Me.OpenFileDialog1.Title = "Укажите ссылку на таблицу"
        Dim dr As DialogResult = Me.OpenFileDialog1.ShowDialog()
        If (dr = System.Windows.Forms.DialogResult.OK) Then
            ' Read the files
            Dim file, fileformat As String
            file = OpenFileDialog1.FileName
            fileformat = file.Substring(file.LastIndexOf(".") + 1)
            If UCase(fileformat) = UCase("xlsx") Or UCase(fileformat) = UCase("xls") Then
                TextBox1.Text = file
                file = ExcelFullFileName
            Else
                MsgBox("Таблица должна быть в формате .xlsx или .xls")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fb As New FolderBrowserDialog
        With fb
            .Description = "Выберите папку с чертежами"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                TextBox2.Text = .SelectedPath
            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            GroupBox1.Enabled = True
        Else
            GroupBox1.Enabled = False
        End If
        butt_enable()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub new_doc_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Arch_ID = Options.Arch_ID
        TextBox1.Text = ExcelFullFileName
        TextBox3.Text = Form1.query_s4("ARCHIVES", "ARCHIVE_ID", "DESCRIPTIO", Arch_ID)
        Button4.Enabled = False
        TextBox2.Text = Form1.Default_work_path
        combobox_add(ComboBox1)
        combobox_add(ComboBox2)
    End Sub
    Sub combobox_add(combobx As ComboBox)
        combobx.Items.Clear()
        For Each txt As String In Form1.doctypes_array.Split("|")
            combobx.Items.Add(txt)
        Next
    End Sub
    Public Sub butt_enable()
        If TextBox1.Text Is Nothing Or TextBox1.Text = "" Then
            Button4.Enabled = False
        Else
            If RadioButton1.Checked Then
                If TextBox2.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Or Arch_ID = 0 Then
                    Button4.Enabled = False
                Else
                    Button4.Enabled = True
                End If
            Else
                Button4.Enabled = True
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        butt_enable()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        butt_enable()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        butt_enable()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        butt_enable()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        butt_enable()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form1.VSNRM_path = TextBox1.Text
        Form1.Doc_list_path = TextBox2.Text
        Form1.detal_doc_types = Form1.query_s4("doctypes", "DOC_NAME", "DOC_TYPE", ComboBox1.Text) ' ComboBox1.Text
        Form1.detal_doc_types_file_Format = Form1.query_s4("doctypes", "DOC_NAME", "DOC_EXT", ComboBox1.Text)
        Form1.SB_doc_types = Form1.query_s4("doctypes", "DOC_NAME", "DOC_TYPE", ComboBox2.Text) ' ComboBox2.Text
        Form1.SB_doc_types_file_Format = Form1.query_s4("doctypes", "DOC_NAME", "DOC_EXT", ComboBox2.Text) ' ComboBox2.Text

        Form1.only_Articles = RadioButton2.Checked
        'If RadioButton2.Checked Then
        '    select_archiv()
        'End If
        Form1.Arch_ID = Arch_ID
        Me.Close()
        Timer1.Interval = 3000
        Application.DoEvents()
        'искусственная задержка на 1 секунду, чтобы успела закрыться окно new_doc НЕ ПОМОГЛО select * from doctypes where doc_ext = 'pdf' And dt_name = 'сборочный чертеж'
        'Thread.Sleep(3000)
        Form1.start_create_BOM()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        select_archiv()
    End Sub
    Function select_archiv()
        Arch_ID = Form1.ITS4App.ChooseArchiveName()
        If Arch_ID <> 0 Then
            TextBox3.Text = Form1.query_s4("ARCHIVES", "ARCHIVE_ID", "DESCRIPTIO", Arch_ID)
        Else
            Exit Function
        End If
    End Function

    Private Sub ОткрытьФайлToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьФайлToolStripMenuItem.Click
        Try
            Open_Second_EX_Doc(True, TextBox1.Text)
        Catch ex As Exception
            MessageBox.Show("Не удается открыть файл. Код ошибки: " & ex.Message)
        End Try
    End Sub

    Private Sub ОткрытьПапкуToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьПапкуToolStripMenuItem.Click
        Try
            Process.Start(TextBox2.Text)
        Catch ex As Exception
            MessageBox.Show("Не удается открыть папку. Код ошибки: " & ex.Message)
        End Try
    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        If TextBox2.Text = "" Then ОткрытьПапкуToolStripMenuItem.Enabled = False Else ОткрытьПапкуToolStripMenuItem.Enabled = True
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If TextBox1.Text = "" Then ОткрытьФайлToolStripMenuItem.Enabled = False Else ОткрытьФайлToolStripMenuItem.Enabled = True
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub new_doc_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.KeyPreview = False
            Me.Close()
        End If
        Me.KeyPreview = True
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        With Form1.ITS4App
            .OpenArchive(Arch_ID)
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Process.Start("http://www.evernote.com/l/AjIzpLy6QAZG35YaM6BgM91l9kUGDbsu-Ak/")
    End Sub
End Class