Public Class Compare_BOM
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
                Form1.progress_txt_im_ListBox1("Подготовка процедуры Сравнения составов")
                Form1.progress_txt_im_ListBox1("Выбран файл ВСНРМ: " & file)
            Else
                MsgBox("Таблица должна быть в формате .xlsx или .xls")
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sel_artID As Integer = Form1.SelectArticl
        If sel_artID > 0 Then
            TextBox2.Text = sel_artID
            Form1.ITS4App.OpenArticle(sel_artID)
            Form1.progress_txt_im_ListBox1("Объект сравнения ArtID: " & sel_artID & " * " & Form1.ITS4App.GetFieldValue_Articles("Обозначение") & " - " & Form1.ITS4App.GetFieldValue_Articles("Наименование"))
            Form1.ITS4App.CloseArticle()
        End If
    End Sub
End Class