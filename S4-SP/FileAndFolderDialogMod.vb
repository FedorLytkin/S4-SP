Module FileAndFolderDialogMod
    Public Function folderBrowserdialog(comments As String) As String
        Dim fb As New FolderBrowserDialog
        With fb
            .Description = "Выберите папку с чертежами"
            If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                folderBrowserdialog = .SelectedPath
            Else
                Exit Function
            End If
        End With
    End Function
    Public Function InitializeOpenFileDialog(Filter As String, Multiselect As Boolean) As String
        Dim OFD As New OpenFileDialog
        'OFD.Filter = "Таблицы(*.xlsx)|"
        OFD.Filter = Filter

        OFD.Multiselect = Multiselect
        OFD.Title = "Укажите ссылку на таблицу"
        Dim file As String
        Dim dr As DialogResult = OFD.ShowDialog()
        If (dr = System.Windows.Forms.DialogResult.OK) Then
            file = OFD.FileName
        End If
        Return file
    End Function
End Module
