Public Class App_info
    Private Sub App_info_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim versionNumber As String

        versionNumber = System.Windows.Forms.Application.ProductVersion
        Me.Label4.Text = "Версия: " & versionNumber
    End Sub

    Private Sub App_info_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then Close()
    End Sub
End Class