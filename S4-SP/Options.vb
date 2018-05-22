Imports System.IO
Imports System.Threading
Public Class Options
    Public material_table_path As String
    Public Search_Subdirectories As String
    Public doctypes_array As String
    Public doctypes_array_default As String = "Сборочный чертеж (TIF)|Чертеж детали (TIF)|Сборочный чертеж (PDF)|Чертеж детали (PDF)"
    Public note_list_check_FFm As String = "изм.; зам."
    Public Default_work_path As String = "C:\IM\IMWork"
    Public excel_visible As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Options_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        doctypes_array = get_reesrt_value("doctypes_array", doctypes_array_default)
        listbox2AddItems()

        material_table_path = get_reesrt_value("material_table_path", Application.StartupPath & "\material_with_ImBaseKey.xlsx")
        Search_Subdirectories = get_reesrt_value("Search_Subdirectories", True.ToString)
        Default_work_path = get_reesrt_value("Default_work_path", Default_work_path)
        excel_visible = get_reesrt_value("excel_visible", False.ToString)

        CheckBox1.Checked = Convert.ToBoolean(Search_Subdirectories)
        TextBox1.Text = material_table_path
        note_list_check_FFm = get_reesrt_value("note_list_check_FFm", note_list_check_FFm)
        TextBox2.Text = note_list_check_FFm
        TextBox4.Text = Default_work_path
        CheckBox2.Checked = Convert.ToBoolean(excel_visible)
    End Sub
    Sub listbox2AddItems()
        ListBox2.Items.Clear()
        For Each txt As String In doctypes_array.Split("|")
            ListBox2.Items.Add(txt)
        Next
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        InitializeOpenFileDialog()
    End Sub
    Public Sub InitializeOpenFileDialog()
        ' Set the file dialog to filter for graphics files.
        Me.OpenFileDialog1.Filter = "Таблицы(*.xlsx)|"


        ' Allow the user to select multiple images.
        Me.OpenFileDialog1.Multiselect = False
        Me.OpenFileDialog1.Title = "Укажите ссылку на таблицу"
        OpenFileDialog1.InitialDirectory = Path.GetDirectoryName(material_table_path)
        Dim dr As DialogResult = Me.OpenFileDialog1.ShowDialog()
        If (dr = System.Windows.Forms.DialogResult.OK) Then
            ' Read the files
            Dim file, fileformat As String
            file = OpenFileDialog1.FileName
            fileformat = file.Substring(file.LastIndexOf(".") + 1)
            If UCase(fileformat) = UCase("xlsx") Or UCase(fileformat) = UCase("xls") Then
                TextBox1.Text = file
            Else
                MsgBox("Таблица должна быть в формате .xlsx или .xls")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        save_options()
        Me.Close()
    End Sub
    Sub save_options()
        set_reesrt_value("material_table_path", TextBox1.Text)
        set_reesrt_value("Search_Subdirectories", CheckBox1.Checked.ToString)
        set_reesrt_value("doctypes_array", doctypes_array_chanche_save)
        note_list_check_FFm = get_reesrt_value("note_list_check_FFm", TextBox2.Text)
        set_reesrt_value("Default_work_path", TextBox4.Text)
        set_reesrt_value("excel_visible", CheckBox2.Checked.ToString)

        Default_work_path = get_reesrt_value("Default_work_path", Default_work_path)

        Form1.material_table_path = TextBox1.Text
        Form1.Search_Subdirectories = CheckBox1.Checked
        Form1.doctypes_array = doctypes_array
        Form1.note_list_check_FFm = note_list_check_FFm
        Form1.Default_work_path = Default_work_path
        Form1.excel_visible = excel_visible
    End Sub
    Function doctypes_array_chanche_save()
        doctypes_array = ""
        For Each txt As String In ListBox2.Items
            If doctypes_array = "" Then doctypes_array = txt Else doctypes_array = doctypes_array & "|" & txt
        Next
        Return doctypes_array
    End Function
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        add_in_User_DocTypes_List()
    End Sub
    Sub add_in_User_DocTypes_List()
        Dim check_exict As Boolean = False
        For Each txt As String In ListBox2.Items
            If txt Is ListBox1.SelectedItem Then check_exict = True
        Next
        If check_exict = False Then ListBox2.Items.Add(ListBox1.SelectedItem)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        remove_in_User_DocTypes_List()
    End Sub
    Sub remove_in_User_DocTypes_List()
        ListBox2.Items.Remove(ListBox2.SelectedItem)
    End Sub
    Private Sub ListBox1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseDoubleClick
        add_in_User_DocTypes_List()
    End Sub

    Private Sub ListBox2_MouseDoubleClick_1(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseDoubleClick
        remove_in_User_DocTypes_List()
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.Items.Count = 0 Then Button1.Enabled = False Else Button1.Enabled = True
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If TextBox1.Text = "" Then ОткрытьФайлToolStripMenuItem.Enabled = False Else ОткрытьФайлToolStripMenuItem.Enabled = True
    End Sub

    Private Sub ОткрытьФайлToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОткрытьФайлToolStripMenuItem.Click
        Try
            Open_Second_EX_Doc(True, TextBox1.Text)
        Catch ex As Exception
            MessageBox.Show("Не удается открыть файл. Код ошибки: " & ex.Message)
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim fb As New FolderBrowserDialog
        With fb
            .Description = "Выберите папку с чертежами"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                TextBox4.Text = .SelectedPath
            Else
                Exit Sub
            End If
        End With
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Try
            Process.Start(TextBox4.Text)
        Catch ex As Exception
            MessageBox.Show("Не удается открыть папку. Код ошибки: " & ex.Message)
        End Try
    End Sub
End Class