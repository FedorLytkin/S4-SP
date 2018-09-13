Public Class Drag_Drop_Form
    Public VSNRM_path As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        Form1.Open_VSNRM(VSNRM_path)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        new_doc.TextBox1.Text = VSNRM_path
        new_doc.ShowDialog()
    End Sub

    Private Sub Drag_Drop_Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.KeyPreview = False
            Me.Close()
        End If
        Me.KeyPreview = True
    End Sub
End Class